from __future__ import annotations

import logging
import os
import json
import re
import socket
from contextlib import contextmanager
from copy import deepcopy
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable

from ledgerflow.config import Config
from ledgerflow.extraction import ExtractionStats, extract_from_files_with_stats
from ledgerflow.google_sheets import (
    GoogleSheetsConfigError,
    SheetsGateway,
    WorksheetNotFoundError,
    build_sheets_gateway,
)
from ledgerflow.google_year_writer import WriteResult, write_yearly_sheet
from ledgerflow.models import Transaction
from ledgerflow.pipeline import PipelineResult, run_pipeline_from_config
from ledgerflow.seed_updater import extract_gpt_transactions

logger = logging.getLogger(__name__)

StatusCallback = Callable[[str], None]
LOCK_CELL_A1 = "M1"
LOCK_TOKEN_PATTERN = re.compile(r"^LOCK\|[^|]+\|\d+\|\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$")
RUN_LOG_HEADERS = [
    "Timestamp",
    "ComputerName",
    "Month",
    "TotalExtracted",
    "AfterDedup",
    "AfterMonthFilter",
    "SeedHits",
    "CacheHits",
    "GptHits",
    "FallbackCount",
    "Status",
    "ErrorMessage",
]


@dataclass(frozen=True)
class DesktopRunSummary:
    extracted_count: int
    deduplicated_count: int
    filtered_count: int
    invalid_date_skips: int
    seed_hits: int
    cache_hits: int
    gpt_hits: int
    fallback_count: int
    backup_path: str
    raw_csv_path: Path
    categorized_csv_path: Path
    final_output_workbook_path: str
    gpt_transactions: list[Transaction]  # transactions resolved by GPT; candidates for seed review


def run_desktop_flow(
    bank_files: list[Path],
    seed_file: Path,
    yearly_workbook: Path,
    month: int,
    cache_file: Path | None = None,
    status_callback: StatusCallback | None = None,
    base_config: Config | None = None,
) -> DesktopRunSummary:
    _validate_inputs(bank_files, seed_file, yearly_workbook, month, cache_file)

    config = _build_runtime_config(seed_file, yearly_workbook, cache_file, base_config)
    output_paths = _build_output_paths()
    gpt_client = _build_openai_client(status_callback)
    gateway = _preflight_google_sheets(config, month, status_callback)
    seed_spreadsheet_id, seed_worksheet_name = config.resolved_seed_target()
    year_spreadsheet_id, year_worksheet_name = config.resolved_year_target(month)
    extraction_stats = ExtractionStats(
        extracted_count=0,
        header_rows_detected=0,
        candidate_rows_seen=0,
        skipped_invalid_date=0,
        skipped_invalid_amount=0,
        skipped_missing_merchant=0,
        skipped_short_rows=0,
    )
    pipeline_result = PipelineResult(
        transactions=[],
        extracted_count=0,
        after_dedup_count=0,
        after_month_filter_count=0,
        seed_count=0,
        cache_count=0,
        gpt_count=0,
        fallback_count=0,
        unresolved_merchants=[],
        cache_was_updated=False,
    )
    run_error = ""
    summary: DesktopRunSummary | None = None

    try:
        with _run_lock(config, month, gateway, status_callback):
            _emit_status(status_callback, "Extracting transactions from bank files...")
            transactions, extraction_stats = extract_from_files_with_stats(bank_files, config)
            _emit_status(
                status_callback,
                (
                    f"Extraction complete: extracted={extraction_stats.extracted_count}, "
                    f"headers={extraction_stats.header_rows_detected}, "
                    f"candidate_rows={extraction_stats.candidate_rows_seen}, "
                    f"skipped_invalid_date={extraction_stats.skipped_invalid_date}, "
                    f"skipped_invalid_amount={extraction_stats.skipped_invalid_amount}, "
                    f"skipped_missing_merchant={extraction_stats.skipped_missing_merchant}, "
                    f"skipped_short_rows={extraction_stats.skipped_short_rows}."
                ),
            )

            _emit_status(status_callback, "Running categorization pipeline...")
            _emit_status(
                status_callback,
                f"Resolved seed target: spreadsheet_id={seed_spreadsheet_id}, worksheet_name={seed_worksheet_name}",
            )
            pipeline_result = run_pipeline_from_config(
                transactions=transactions,
                month=month,
                config=config,
                gpt_client=gpt_client,
                raw_csv_path=output_paths.raw_csv_path,
                categorized_csv_path=output_paths.categorized_csv_path,
            )
            _emit_status(
                status_callback,
                (
                    f"Pipeline counts: extracted={pipeline_result.extracted_count}, "
                    f"after_dedup={pipeline_result.after_dedup_count}, "
                    f"after_month_filter={pipeline_result.after_month_filter_count}."
                ),
            )

            if pipeline_result.after_month_filter_count == 0:
                raise ValueError(
                    f"No transactions were found for month {month}. "
                    f"Counts: extracted={pipeline_result.extracted_count}, "
                    f"after_dedup={pipeline_result.after_dedup_count}, "
                    f"after_month_filter={pipeline_result.after_month_filter_count}, "
                    f"skipped_invalid_date={extraction_stats.skipped_invalid_date}. "
                    "Choose a different month or check the selected files."
                )

            _emit_status(status_callback, "Writing yearly Google Sheet output...")
            _emit_status(
                status_callback,
                f"Resolved year target: spreadsheet_id={year_spreadsheet_id}, worksheet_name={year_worksheet_name}",
            )
            write_result = write_yearly_sheet(
                transactions=pipeline_result.transactions,
                spreadsheet_id=year_spreadsheet_id,
                month=month,
                config=config,
                gateway=gateway,
            )

        _emit_status(status_callback, "Run completed successfully.")
        summary = _build_summary(
            pipeline_result, write_result, output_paths, extraction_stats
        )
        return summary
    except Exception as exc:
        run_error = str(exc)
        raise
    finally:
        _append_run_log(
            config=config,
            month=month,
            gateway=gateway,
            extraction_stats=extraction_stats,
            pipeline_result=pipeline_result,
            error_message=run_error,
        )


@dataclass(frozen=True)
class _OutputPaths:
    output_dir: Path
    raw_csv_path: Path
    categorized_csv_path: Path
    final_output_workbook_path: Path


def _build_summary(
    pipeline_result: PipelineResult,
    write_result: WriteResult,
    output_paths: _OutputPaths,
    extraction_stats: ExtractionStats,
) -> DesktopRunSummary:
    non_gpt_cache_hits = max(pipeline_result.cache_count - pipeline_result.gpt_count, 0)
    return DesktopRunSummary(
        extracted_count=pipeline_result.extracted_count,
        deduplicated_count=pipeline_result.after_dedup_count,
        filtered_count=pipeline_result.after_month_filter_count,
        invalid_date_skips=extraction_stats.skipped_invalid_date,
        seed_hits=pipeline_result.seed_count,
        cache_hits=non_gpt_cache_hits,
        gpt_hits=pipeline_result.gpt_count,
        fallback_count=pipeline_result.fallback_count,
        backup_path=write_result.backup_path,
        raw_csv_path=output_paths.raw_csv_path,
        categorized_csv_path=output_paths.categorized_csv_path,
        final_output_workbook_path=write_result.output_path,
        gpt_transactions=extract_gpt_transactions(pipeline_result.transactions),
    )


def _build_runtime_config(
    seed_file: Path,
    yearly_workbook: Path,
    cache_file: Path | None,
    base_config: Config | None = None,
) -> Config:
    cfg = deepcopy(base_config) if base_config is not None else Config()
    cfg.cache_path = cache_file
    cfg.seed_spreadsheet_id = str(seed_file)
    cfg.year_spreadsheet_id = str(yearly_workbook)
    seed_worksheet = os.getenv("LEDGERFLOW_SEED_WORKSHEET", "").strip()
    if seed_worksheet:
        cfg.seed_worksheet_name = seed_worksheet
    year_worksheet_names = os.getenv("LEDGERFLOW_YEAR_WORKSHEET_NAMES", "").strip()
    if year_worksheet_names:
        cfg.year_worksheet_names = {
            int(month): str(name)
            for month, name in json.loads(year_worksheet_names).items()
        }
    run_log_worksheet = os.getenv("LEDGERFLOW_RUNLOG_WORKSHEET", "").strip()
    if run_log_worksheet:
        cfg.run_log_worksheet_name = run_log_worksheet
    service_account = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "").strip()
    if service_account:
        cfg.google_service_account_file = Path(service_account)
    return cfg


def _build_output_paths() -> _OutputPaths:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = Path.cwd() / f"ledgerflow_run_{timestamp}"
    output_dir.mkdir(parents=True, exist_ok=True)

    return _OutputPaths(
        output_dir=output_dir,
        raw_csv_path=output_dir / "raw_transactions.csv",
        categorized_csv_path=output_dir / "categorized_transactions.csv",
        final_output_workbook_path=output_dir / f"year_sheet_{timestamp}.txt",
    )


def _build_openai_client(status_callback: StatusCallback | None) -> object | None:
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        _emit_status(
            status_callback,
            "OPENAI_API_KEY not set. GPT categorization is disabled for this run.",
        )
        return None

    try:
        from openai import OpenAI
    except ImportError:
        _emit_status(
            status_callback,
            "OpenAI package is not installed. GPT categorization is disabled for this run.",
        )
        return None

    _emit_status(status_callback, "GPT categorization is enabled for unresolved merchants.")
    return OpenAI(api_key=api_key)


def _preflight_google_sheets(
    config: Config,
    month: int,
    status_callback: StatusCallback | None,
) -> SheetsGateway:
    _emit_status(status_callback, "Validating Google Sheets access...")
    try:
        seed_spreadsheet_id, seed_worksheet_name = config.resolved_seed_target()
        year_spreadsheet_id, year_worksheet_name = config.resolved_year_target(month)
    except Exception as exc:  # noqa: BLE001
        raise _google_sheets_failure(
            "config resolution",
            exc,
            status_callback,
        ) from exc

    try:
        gateway = build_sheets_gateway(config.google_service_account_file)
    except Exception as exc:  # noqa: BLE001
        raise _google_sheets_failure(
            "auth init",
            exc,
            status_callback,
            details=f"credentials_path={config.google_service_account_file!s}",
        ) from exc

    try:
        gateway.read_sheet(seed_spreadsheet_id, seed_worksheet_name)
    except Exception as exc:  # noqa: BLE001
        raise _google_sheets_failure(
            "seed worksheet lookup",
            exc,
            status_callback,
            details=(
                f"spreadsheet_id={seed_spreadsheet_id}, "
                f"worksheet_name={seed_worksheet_name}"
            ),
        ) from exc

    try:
        year_sheet_names = gateway.list_sheet_names(year_spreadsheet_id)
    except Exception as exc:  # noqa: BLE001
        raise _google_sheets_failure(
            "year spreadsheet open",
            exc,
            status_callback,
            details=f"spreadsheet_id={year_spreadsheet_id}",
        ) from exc

    if year_worksheet_name not in year_sheet_names:
        exc = ValueError(
            f"Configured year worksheet {year_worksheet_name!r} for month {month} was not found."
        )
        raise _google_sheets_failure(
            "year month worksheet lookup",
            exc,
            status_callback,
            details=(
                f"spreadsheet_id={year_spreadsheet_id}, "
                f"worksheet_name={year_worksheet_name}"
            ),
        ) from exc

    try:
        gateway.read_sheet(year_spreadsheet_id, year_worksheet_name)
    except Exception as exc:  # noqa: BLE001
        raise _google_sheets_failure(
            "year month worksheet lookup",
            exc,
            status_callback,
            details=(
                f"spreadsheet_id={year_spreadsheet_id}, "
                f"worksheet_name={year_worksheet_name}"
            ),
        ) from exc

    if config.run_log_worksheet_name not in year_sheet_names:
        exc = ValueError(
            f"Run log worksheet {config.run_log_worksheet_name!r} was not found in the year spreadsheet."
        )
        raise _google_sheets_failure(
            "run log worksheet lookup",
            exc,
            status_callback,
            details=(
                f"spreadsheet_id={year_spreadsheet_id}, "
                f"worksheet_name={config.run_log_worksheet_name}"
            ),
        ) from exc

    _emit_status(status_callback, "Google Sheets validation passed.")
    return gateway


@contextmanager
def _run_lock(
    config: Config,
    month: int,
    gateway: SheetsGateway,
    status_callback: StatusCallback | None,
):
    year_spreadsheet_id = config.year_spreadsheet_id
    if not year_spreadsheet_id:
        raise ValueError("year.spreadsheet_id is not configured.")
    worksheet_name = config.run_log_worksheet_name
    token = _lock_token()
    try:
        existing = _read_lock_value(gateway, year_spreadsheet_id, worksheet_name)
    except Exception as exc:  # noqa: BLE001
        raise _google_sheets_failure(
            "lock read",
            exc,
            status_callback,
            details=(
                f"spreadsheet_id={year_spreadsheet_id}, "
                f"worksheet_name={worksheet_name}, range={LOCK_CELL_A1}"
            ),
        ) from exc
    if existing:
        raise ValueError(
            f"Another LedgerFlow run is already in progress: {existing}. "
            "Wait for it to finish and try again."
        )

    try:
        gateway.update_range(
            year_spreadsheet_id,
            worksheet_name,
            LOCK_CELL_A1,
            [[token]],
        )
        confirmed = _read_lock_value(gateway, year_spreadsheet_id, worksheet_name)
    except Exception as exc:  # noqa: BLE001
        raise _google_sheets_failure(
            "lock acquire",
            exc,
            status_callback,
            details=(
                f"spreadsheet_id={year_spreadsheet_id}, "
                f"worksheet_name={worksheet_name}, range={LOCK_CELL_A1}"
            ),
        ) from exc
    if confirmed != token:
        raise ValueError("Could not acquire the shared LedgerFlow run lock.")

    _emit_status(status_callback, "Run lock acquired.")
    try:
        yield
    finally:
        try:
            current = _read_lock_value(gateway, year_spreadsheet_id, worksheet_name)
            if current == token:
                gateway.update_range(
                    year_spreadsheet_id,
                    worksheet_name,
                    LOCK_CELL_A1,
                    [[""]],
                )
        except Exception as exc:  # noqa: BLE001
            _emit_status(
                status_callback,
                _google_sheets_failure_message(
                    "lock release",
                    exc,
                    (
                        f"spreadsheet_id={year_spreadsheet_id}, "
                        f"worksheet_name={worksheet_name}, range={LOCK_CELL_A1}"
                    ),
                ),
            )
        _emit_status(status_callback, "Run lock released.")


def _read_lock_value(
    gateway: SheetsGateway,
    spreadsheet_id: str,
    worksheet_name: str,
) -> str:
    values = gateway.read_range(spreadsheet_id, worksheet_name, LOCK_CELL_A1)
    if not values or not values[0]:
        return ""
    value = values[0][0]
    text = str(value).strip() if value is not None else ""
    return text if LOCK_TOKEN_PATTERN.fullmatch(text) else ""


def _lock_token() -> str:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return f"LOCK|{socket.gethostname()}|{os.getpid()}|{timestamp}"


def _append_run_log(
    config: Config,
    month: int,
    gateway: SheetsGateway,
    extraction_stats: ExtractionStats,
    pipeline_result: PipelineResult,
    error_message: str,
) -> None:
    if config.year_spreadsheet_id is None:
        return
    cache_hits = max(pipeline_result.cache_count - pipeline_result.gpt_count, 0)
    row = [
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        socket.gethostname(),
        month,
        pipeline_result.extracted_count or extraction_stats.extracted_count,
        pipeline_result.after_dedup_count,
        pipeline_result.after_month_filter_count,
        pipeline_result.seed_count,
        cache_hits,
        pipeline_result.gpt_count,
        pipeline_result.fallback_count,
        "failure" if error_message else "success",
        error_message,
    ]
    gateway.append_rows(
        config.year_spreadsheet_id,
        config.run_log_worksheet_name,
        [RUN_LOG_HEADERS] if _run_log_is_empty(gateway, config) else [],
    )
    gateway.append_rows(
        config.year_spreadsheet_id,
        config.run_log_worksheet_name,
        [row],
    )


def _run_log_is_empty(gateway: SheetsGateway, config: Config) -> bool:
    if config.year_spreadsheet_id is None:
        return False
    rows = gateway.read_sheet(config.year_spreadsheet_id, config.run_log_worksheet_name)
    return len(rows) == 0


def _validate_inputs(
    bank_files: list[Path],
    seed_file: Path,
    yearly_workbook: Path,
    month: int,
    cache_file: Path | None,
) -> None:
    if not bank_files:
        raise ValueError("Select at least one bank transaction file.")
    if str(seed_file).strip() in {"", "."}:
        raise ValueError("Enter a seed spreadsheet ID.")
    if str(yearly_workbook).strip() in {"", "."}:
        raise ValueError("Enter a year spreadsheet ID.")

    missing_paths = [path for path in bank_files if not path.exists()]
    if cache_file is not None and cache_file.exists() is False and cache_file.parent.exists() is False:
        missing_paths.append(cache_file.parent)

    if missing_paths:
        missing_text = ", ".join(str(path) for path in missing_paths)
        raise FileNotFoundError(f"Required file or folder was not found: {missing_text}")

    if not 1 <= month <= 12:
        raise ValueError(f"Month must be between 1 and 12. Received {month}.")


def _emit_status(status_callback: StatusCallback | None, message: str) -> None:
    logger.info(message)
    if status_callback is not None:
        status_callback(message)


def _google_sheets_failure(
    step: str,
    exc: Exception,
    status_callback: StatusCallback | None,
    details: str | None = None,
) -> ValueError:
    message = _google_sheets_failure_message(step, exc, details)
    _emit_status(status_callback, message)
    return ValueError(message)


def _google_sheets_failure_message(
    step: str,
    exc: Exception,
    details: str | None = None,
) -> str:
    message = f"Google Sheets preflight failed during {step}"
    if details:
        message += f" ({details})"
    message += f": {exc}"
    return message
