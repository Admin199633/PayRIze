from __future__ import annotations

import logging
import os
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable

from ledgerflow.config import Config
from ledgerflow.extraction import ExtractionStats, extract_from_files_with_stats
from ledgerflow.models import Transaction
from ledgerflow.pipeline import PipelineResult, run_pipeline_from_config
from ledgerflow.seed_updater import extract_gpt_transactions
from ledgerflow.year_writer import WriteResult, write_yearly_workbook

logger = logging.getLogger(__name__)

StatusCallback = Callable[[str], None]


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
    backup_path: Path
    raw_csv_path: Path
    categorized_csv_path: Path
    final_output_workbook_path: Path
    gpt_transactions: list[Transaction]  # transactions resolved by GPT; candidates for seed review


def run_desktop_flow(
    bank_files: list[Path],
    seed_file: Path,
    yearly_workbook: Path,
    month: int,
    cache_file: Path | None = None,
    status_callback: StatusCallback | None = None,
) -> DesktopRunSummary:
    _validate_inputs(bank_files, seed_file, yearly_workbook, month, cache_file)

    config = Config(seed_mapping_path=seed_file, cache_path=cache_file)
    output_paths = _build_output_paths(yearly_workbook)
    gpt_client = _build_openai_client(status_callback)

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

    _emit_status(status_callback, "Writing yearly workbook output...")
    write_result = write_yearly_workbook(
        transactions=pipeline_result.transactions,
        source_workbook=yearly_workbook,
        output_path=output_paths.final_output_workbook_path,
        month=month,
        config=config,
        backup_dir=output_paths.output_dir,
    )

    _emit_status(status_callback, "Run completed successfully.")
    return _build_summary(
        pipeline_result, write_result, output_paths, extraction_stats
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


def _build_output_paths(yearly_workbook: Path) -> _OutputPaths:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = yearly_workbook.parent / f"ledgerflow_run_{timestamp}"
    output_dir.mkdir(parents=True, exist_ok=True)

    return _OutputPaths(
        output_dir=output_dir,
        raw_csv_path=output_dir / "raw_transactions.csv",
        categorized_csv_path=output_dir / "categorized_transactions.csv",
        final_output_workbook_path=output_dir
        / f"{yearly_workbook.stem}_processed_{timestamp}{yearly_workbook.suffix}",
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
        raise ValueError("Select a seed file.")
    if str(yearly_workbook).strip() in {"", "."}:
        raise ValueError("Select a yearly workbook file.")

    missing_paths = [path for path in [*bank_files, seed_file, yearly_workbook] if not path.exists()]
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
