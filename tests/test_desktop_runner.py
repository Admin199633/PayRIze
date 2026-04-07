from __future__ import annotations

from pathlib import Path
from unittest.mock import sentinel

from ledgerflow.config import Config
from ledgerflow.desktop_runner import _read_lock_value, run_desktop_flow
from ledgerflow.extraction import ExtractionStats
from ledgerflow.google_year_writer import WriteResult
from ledgerflow.pipeline import PipelineResult


class FakeSheetsGateway:
    def __init__(self, lock_value: str = "") -> None:
        self.seed_rows = [["MerchantKey", "MerchantOriginal", "Category"]]
        self.year_rows = [["Other", "Price", "Date"]]
        self.run_log_rows: list[list[object]] = []
        self.lock_value = lock_value

    def list_sheet_names(self, spreadsheet_id: str) -> list[str]:
        return ["ConfiguredMarch", "RunLog"]

    def read_sheet(self, spreadsheet_id: str, worksheet_name: str) -> list[list[object]]:
        if spreadsheet_id == "seed-spreadsheet-id":
            return [list(row) for row in self.seed_rows]
        if spreadsheet_id == "year-spreadsheet-id":
            if worksheet_name == "RunLog":
                return [list(row) for row in self.run_log_rows]
            return [list(row) for row in self.year_rows]
        raise AssertionError(f"Unexpected spreadsheet_id {spreadsheet_id}")

    def batch_update_rows(self, spreadsheet_id: str, worksheet_name: str, updates) -> None:
        return None

    def append_rows(self, spreadsheet_id: str, worksheet_name: str, rows) -> None:
        if spreadsheet_id != "year-spreadsheet-id" or worksheet_name != "RunLog":
            raise AssertionError(f"Unexpected append target {spreadsheet_id}/{worksheet_name}")
        for row in rows:
            self.run_log_rows.append(list(row))

    def read_range(self, spreadsheet_id: str, worksheet_name: str, a1_range: str) -> list[list[object]]:
        return [[self.lock_value]] if self.lock_value else []

    def update_range(self, spreadsheet_id: str, worksheet_name: str, a1_range: str, values) -> None:
        raw = values[0][0] if values and values[0] else ""
        self.lock_value = str(raw).strip()


def test_run_desktop_flow_orchestrates_backend_modules(monkeypatch, tmp_path):
    bank_file = tmp_path / "bank.xlsx"
    bank_file.write_text("x", encoding="utf-8")
    seed_file = Path("seed-spreadsheet-id")
    yearly_workbook = Path("year-spreadsheet-id")

    extracted = [sentinel.transaction]
    statuses: list[str] = []
    captured: dict[str, object] = {}

    def fake_extract(paths, config):
        captured["extract_paths"] = paths
        captured["extract_config"] = config
        return extracted, ExtractionStats(
            extracted_count=1,
            header_rows_detected=1,
            candidate_rows_seen=1,
            skipped_invalid_date=0,
            skipped_invalid_amount=0,
            skipped_missing_merchant=0,
            skipped_short_rows=0,
        )

    def fake_pipeline(*, transactions, month, config, gpt_client, raw_csv_path, categorized_csv_path):
        captured["pipeline_transactions"] = transactions
        captured["pipeline_month"] = month
        captured["pipeline_config"] = config
        captured["pipeline_gpt_client"] = gpt_client
        captured["raw_csv_path"] = raw_csv_path
        captured["categorized_csv_path"] = categorized_csv_path
        return PipelineResult(
            transactions=[sentinel.categorized_transaction],
            extracted_count=14,
            after_dedup_count=11,
            after_month_filter_count=7,
            seed_count=3,
            cache_count=4,
            gpt_count=1,
            fallback_count=0,
            unresolved_merchants=[],
            cache_was_updated=True,
        )

    def fake_write(*, transactions, spreadsheet_id, month, config, gateway):
        captured["write_transactions"] = transactions
        captured["write_spreadsheet_id"] = spreadsheet_id
        captured["write_month"] = month
        captured["write_config"] = config
        captured["write_gateway"] = gateway
        return WriteResult(
            backup_path="not-applicable",
            output_path="https://docs.google.com/spreadsheets/d/year-spreadsheet-id/edit",
            transactions_written=7,
            skipped_count=0,
        )

    monkeypatch.setattr("ledgerflow.desktop_runner.extract_from_files_with_stats", fake_extract)
    monkeypatch.setattr("ledgerflow.desktop_runner.run_pipeline_from_config", fake_pipeline)
    monkeypatch.setattr("ledgerflow.desktop_runner.write_yearly_sheet", fake_write)
    monkeypatch.setattr("ledgerflow.desktop_runner._build_openai_client", lambda cb: sentinel.gpt_client)
    monkeypatch.setattr("ledgerflow.desktop_runner.extract_gpt_transactions", lambda txs: [])
    gateway = FakeSheetsGateway()
    monkeypatch.setattr("ledgerflow.desktop_runner.build_sheets_gateway", lambda service_account_file=None: gateway)
    monkeypatch.setenv("LEDGERFLOW_SEED_WORKSHEET", "Seed")
    monkeypatch.setenv("LEDGERFLOW_YEAR_WORKSHEET_NAMES", '{"3":"ConfiguredMarch"}')

    summary = run_desktop_flow(
        bank_files=[bank_file],
        seed_file=seed_file,
        yearly_workbook=yearly_workbook,
        month=3,
        cache_file=None,
        status_callback=statuses.append,
    )

    assert captured["extract_paths"] == [bank_file]
    assert captured["pipeline_transactions"] == extracted
    assert captured["pipeline_month"] == 3
    assert captured["pipeline_gpt_client"] is sentinel.gpt_client
    assert captured["write_transactions"] == [sentinel.categorized_transaction]
    assert captured["write_spreadsheet_id"] == "year-spreadsheet-id"
    assert captured["write_month"] == 3

    assert summary.extracted_count == 14
    assert summary.deduplicated_count == 11
    assert summary.filtered_count == 7
    assert summary.invalid_date_skips == 0
    assert summary.seed_hits == 3
    assert summary.cache_hits == 3
    assert summary.gpt_hits == 1
    assert summary.fallback_count == 0
    assert summary.raw_csv_path == captured["raw_csv_path"]
    assert summary.categorized_csv_path == captured["categorized_csv_path"]
    assert summary.final_output_workbook_path == "https://docs.google.com/spreadsheets/d/year-spreadsheet-id/edit"
    assert any("Google Sheets validation passed" in status for status in statuses)
    assert any("Run lock acquired" in status for status in statuses)
    assert any("Run lock released" in status for status in statuses)
    assert any("Resolved seed target: spreadsheet_id=seed-spreadsheet-id, worksheet_name=Seed" in status for status in statuses)
    assert any("Resolved year target: spreadsheet_id=year-spreadsheet-id, worksheet_name=ConfiguredMarch" in status for status in statuses)
    assert any("Extracting transactions" in status for status in statuses)
    assert any("Running categorization pipeline" in status for status in statuses)
    assert any("Writing yearly Google Sheet output" in status for status in statuses)
    assert gateway.run_log_rows[0] == [
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
    assert gateway.run_log_rows[1][2] == 3
    assert gateway.run_log_rows[1][3] == 14
    assert gateway.run_log_rows[1][4] == 11
    assert gateway.run_log_rows[1][5] == 7
    assert gateway.run_log_rows[1][6] == 3
    assert gateway.run_log_rows[1][7] == 3
    assert gateway.run_log_rows[1][8] == 1
    assert gateway.run_log_rows[1][9] == 0
    assert gateway.run_log_rows[1][10] == "success"
    assert gateway.run_log_rows[1][11] == ""


def test_run_desktop_flow_rejects_empty_month_result(monkeypatch, tmp_path):
    bank_file = tmp_path / "bank.xlsx"
    bank_file.write_text("x", encoding="utf-8")
    seed_file = Path("seed-spreadsheet-id")
    yearly_workbook = Path("year-spreadsheet-id")

    monkeypatch.setattr(
        "ledgerflow.desktop_runner.extract_from_files_with_stats",
        lambda paths, config: ([], ExtractionStats(
            extracted_count=0,
            header_rows_detected=0,
            candidate_rows_seen=0,
            skipped_invalid_date=2,
            skipped_invalid_amount=0,
            skipped_missing_merchant=0,
            skipped_short_rows=0,
        )),
    )
    monkeypatch.setattr("ledgerflow.desktop_runner._build_openai_client", lambda cb: None)
    gateway = FakeSheetsGateway()
    monkeypatch.setattr("ledgerflow.desktop_runner.build_sheets_gateway", lambda service_account_file=None: gateway)
    monkeypatch.setenv("LEDGERFLOW_YEAR_WORKSHEET_NAMES", '{"3":"ConfiguredMarch"}')

    def fake_pipeline(**kwargs):
        return PipelineResult(
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

    monkeypatch.setattr("ledgerflow.desktop_runner.run_pipeline_from_config", fake_pipeline)

    try:
        run_desktop_flow(
            bank_files=[bank_file],
            seed_file=seed_file,
            yearly_workbook=yearly_workbook,
            month=3,
        )
    except ValueError as exc:
        assert "No transactions were found for month 3" in str(exc)
    else:
        raise AssertionError("Expected ValueError for an empty month result")
    assert gateway.run_log_rows[1][10] == "failure"
    assert "No transactions were found for month 3" in gateway.run_log_rows[1][11]


def test_run_desktop_flow_fails_clearly_when_google_sheets_preflight_fails(monkeypatch, tmp_path):
    bank_file = tmp_path / "bank.xlsx"
    bank_file.write_text("x", encoding="utf-8")
    statuses: list[str] = []

    def fail_gateway(service_account_file=None):
        raise RuntimeError("boom from auth init")

    monkeypatch.setattr("ledgerflow.desktop_runner.build_sheets_gateway", fail_gateway)
    monkeypatch.setattr("ledgerflow.desktop_runner._build_openai_client", lambda cb: None)
    monkeypatch.setenv("LEDGERFLOW_YEAR_WORKSHEET_NAMES", '{"3":"ConfiguredMarch"}')

    try:
        run_desktop_flow(
            bank_files=[bank_file],
            seed_file=Path("seed-spreadsheet-id"),
            yearly_workbook=Path("year-spreadsheet-id"),
            month=3,
            status_callback=statuses.append,
        )
    except ValueError as exc:
        assert "Google Sheets preflight failed during auth init" in str(exc)
        assert "credentials_path=" in str(exc)
        assert "boom from auth init" in str(exc)
        assert "Check service-account auth and spreadsheet sharing" not in str(exc)
    else:
        raise AssertionError("Expected ValueError for Google Sheets preflight failure")
    assert any("Google Sheets preflight failed during auth init" in status for status in statuses)


def test_run_desktop_flow_preserves_preflight_seed_lookup_error_text_and_target(monkeypatch, tmp_path):
    bank_file = tmp_path / "bank.xlsx"
    bank_file.write_text("x", encoding="utf-8")

    class FailingSeedLookupGateway(FakeSheetsGateway):
        def read_sheet(self, spreadsheet_id: str, worksheet_name: str) -> list[list[object]]:
            if spreadsheet_id == "seed-spreadsheet-id" and worksheet_name == "Seed":
                raise RuntimeError("Worksheet 'Seed' was not found")
            return super().read_sheet(spreadsheet_id, worksheet_name)

    monkeypatch.setattr(
        "ledgerflow.desktop_runner.build_sheets_gateway",
        lambda service_account_file=None: FailingSeedLookupGateway(),
    )
    monkeypatch.setattr("ledgerflow.desktop_runner._build_openai_client", lambda cb: None)
    monkeypatch.setenv("LEDGERFLOW_SEED_WORKSHEET", "Seed")
    monkeypatch.setenv("LEDGERFLOW_YEAR_WORKSHEET_NAMES", '{"3":"ConfiguredMarch"}')

    try:
        run_desktop_flow(
            bank_files=[bank_file],
            seed_file=Path("seed-spreadsheet-id"),
            yearly_workbook=Path("year-spreadsheet-id"),
            month=3,
        )
    except ValueError as exc:
        text = str(exc)
        assert "Google Sheets preflight failed during seed worksheet lookup" in text
        assert "spreadsheet_id=seed-spreadsheet-id, worksheet_name=Seed" in text
        assert "Worksheet 'Seed' was not found" in text
    else:
        raise AssertionError("Expected ValueError for seed worksheet lookup failure")


def test_run_desktop_flow_rejects_when_lock_is_already_held(monkeypatch, tmp_path):
    bank_file = tmp_path / "bank.xlsx"
    bank_file.write_text("x", encoding="utf-8")
    monkeypatch.setattr(
        "ledgerflow.desktop_runner.build_sheets_gateway",
        lambda service_account_file=None: FakeSheetsGateway(lock_value="LOCK|other-host|1|2026-04-07 11:00:00"),
    )
    monkeypatch.setattr("ledgerflow.desktop_runner._build_openai_client", lambda cb: None)
    monkeypatch.setenv("LEDGERFLOW_YEAR_WORKSHEET_NAMES", '{"3":"ConfiguredMarch"}')

    try:
        run_desktop_flow(
            bank_files=[bank_file],
            seed_file=Path("seed-spreadsheet-id"),
            yearly_workbook=Path("year-spreadsheet-id"),
            month=3,
        )
    except ValueError as exc:
        assert "already in progress" in str(exc)
    else:
        raise AssertionError("Expected ValueError when the shared run lock is occupied")


def test_run_desktop_flow_releases_lock_on_failure(monkeypatch, tmp_path):
    bank_file = tmp_path / "bank.xlsx"
    bank_file.write_text("x", encoding="utf-8")
    gateway = FakeSheetsGateway()

    monkeypatch.setattr("ledgerflow.desktop_runner.build_sheets_gateway", lambda service_account_file=None: gateway)
    monkeypatch.setattr("ledgerflow.desktop_runner._build_openai_client", lambda cb: None)
    monkeypatch.setenv("LEDGERFLOW_YEAR_WORKSHEET_NAMES", '{"3":"ConfiguredMarch"}')

    def fail_extract(paths, config):
        raise RuntimeError("boom")

    monkeypatch.setattr("ledgerflow.desktop_runner.extract_from_files_with_stats", fail_extract)

    try:
        run_desktop_flow(
            bank_files=[bank_file],
            seed_file=Path("seed-spreadsheet-id"),
            yearly_workbook=Path("year-spreadsheet-id"),
            month=3,
        )
    except RuntimeError:
        pass
    else:
        raise AssertionError("Expected RuntimeError from extraction failure")

    assert gateway.lock_value == ""
    assert gateway.run_log_rows[1][10] == "failure"
    assert gateway.run_log_rows[1][11] == "boom"
    assert gateway.lock_value == ""


def test_run_desktop_flow_fails_clearly_when_runlog_worksheet_missing(monkeypatch, tmp_path):
    bank_file = tmp_path / "bank.xlsx"
    bank_file.write_text("x", encoding="utf-8")

    class MissingRunLogGateway(FakeSheetsGateway):
        def list_sheet_names(self, spreadsheet_id: str) -> list[str]:
            return ["March"]

    monkeypatch.setattr(
        "ledgerflow.desktop_runner.build_sheets_gateway",
        lambda service_account_file=None: MissingRunLogGateway(),
    )
    monkeypatch.setattr("ledgerflow.desktop_runner._build_openai_client", lambda cb: None)
    monkeypatch.setenv("LEDGERFLOW_YEAR_WORKSHEET_NAMES", '{"3":"ConfiguredMarch"}')

    try:
        run_desktop_flow(
            bank_files=[bank_file],
            seed_file=Path("seed-spreadsheet-id"),
            yearly_workbook=Path("year-spreadsheet-id"),
            month=3,
        )
    except ValueError as exc:
        assert "Google Sheets preflight failed during run log worksheet lookup" in str(exc)
        assert "worksheet_name=RunLog" in str(exc)
        assert "Run log worksheet 'RunLog' was not found in the year spreadsheet." in str(exc)
    else:
        raise AssertionError("Expected ValueError when RunLog worksheet is missing")


def test_run_desktop_flow_fails_clearly_when_month_mapping_missing(monkeypatch, tmp_path):
    bank_file = tmp_path / "bank.xlsx"
    bank_file.write_text("x", encoding="utf-8")

    monkeypatch.setattr(
        "ledgerflow.desktop_runner.build_sheets_gateway",
        lambda service_account_file=None: FakeSheetsGateway(),
    )
    monkeypatch.setattr("ledgerflow.desktop_runner._build_openai_client", lambda cb: None)

    try:
        run_desktop_flow(
            bank_files=[bank_file],
            seed_file=Path("seed-spreadsheet-id"),
            yearly_workbook=Path("year-spreadsheet-id"),
            month=3,
        )
    except ValueError as exc:
        assert "year.month_sheets does not contain an entry for month 3" in str(exc)
    else:
        raise AssertionError("Expected ValueError when month_sheets has no configured target")


def test_run_desktop_flow_preserves_loaded_month_mapping_from_base_config(monkeypatch, tmp_path):
    bank_file = tmp_path / "bank.xlsx"
    bank_file.write_text("x", encoding="utf-8")

    extracted = [sentinel.transaction]
    captured: dict[str, object] = {}

    def fake_extract(paths, config):
        captured["extract_config"] = config
        return extracted, ExtractionStats(
            extracted_count=1,
            header_rows_detected=1,
            candidate_rows_seen=1,
            skipped_invalid_date=0,
            skipped_invalid_amount=0,
            skipped_missing_merchant=0,
            skipped_short_rows=0,
        )

    def fake_pipeline(*, transactions, month, config, gpt_client, raw_csv_path, categorized_csv_path):
        captured["pipeline_config"] = config
        return PipelineResult(
            transactions=[sentinel.categorized_transaction],
            extracted_count=1,
            after_dedup_count=1,
            after_month_filter_count=1,
            seed_count=0,
            cache_count=0,
            gpt_count=0,
            fallback_count=1,
            unresolved_merchants=[],
            cache_was_updated=False,
        )

    def fake_write(*, transactions, spreadsheet_id, month, config, gateway):
        captured["write_config"] = config
        return WriteResult(
            backup_path="not-applicable",
            output_path="https://docs.google.com/spreadsheets/d/year-spreadsheet-id/edit",
            transactions_written=1,
            skipped_count=0,
        )

    base_config = Config()
    base_config.seed_spreadsheet_id = "seed-from-config"
    base_config.year_spreadsheet_id = "year-from-config"
    base_config.year_worksheet_names = {3: "ConfiguredMarch"}

    monkeypatch.setattr("ledgerflow.desktop_runner.extract_from_files_with_stats", fake_extract)
    monkeypatch.setattr("ledgerflow.desktop_runner.run_pipeline_from_config", fake_pipeline)
    monkeypatch.setattr("ledgerflow.desktop_runner.write_yearly_sheet", fake_write)
    monkeypatch.setattr("ledgerflow.desktop_runner._build_openai_client", lambda cb: None)
    monkeypatch.setattr("ledgerflow.desktop_runner.extract_gpt_transactions", lambda txs: [])
    monkeypatch.setattr("ledgerflow.desktop_runner.build_sheets_gateway", lambda service_account_file=None: FakeSheetsGateway())

    run_desktop_flow(
        bank_files=[bank_file],
        seed_file=Path("seed-manual"),
        yearly_workbook=Path("year-manual"),
        month=3,
        base_config=base_config,
    )

    for key in ("extract_config", "pipeline_config", "write_config"):
        cfg = captured[key]
        assert cfg.seed_spreadsheet_id == "seed-manual"
        assert cfg.year_spreadsheet_id == "year-manual"
        assert cfg.year_worksheet_names == {3: "ConfiguredMarch"}


def test_empty_lock_cell_is_not_treated_as_active_lock(monkeypatch, tmp_path):
    bank_file = tmp_path / "bank.xlsx"
    bank_file.write_text("x", encoding="utf-8")
    gateway = FakeSheetsGateway(lock_value="")

    monkeypatch.setattr("ledgerflow.desktop_runner.build_sheets_gateway", lambda service_account_file=None: gateway)
    monkeypatch.setattr("ledgerflow.desktop_runner._build_openai_client", lambda cb: None)
    monkeypatch.setenv("LEDGERFLOW_YEAR_WORKSHEET_NAMES", '{"3":"ConfiguredMarch"}')
    monkeypatch.setattr(
        "ledgerflow.desktop_runner.extract_from_files_with_stats",
        lambda paths, config: ([], ExtractionStats()),
    )
    monkeypatch.setattr(
        "ledgerflow.desktop_runner.run_pipeline_from_config",
        lambda **kwargs: PipelineResult([], 0, 0, 0, 0, 0, 0, 0, [], False),
    )

    try:
        run_desktop_flow(
            bank_files=[bank_file],
            seed_file=Path("seed-spreadsheet-id"),
            yearly_workbook=Path("year-spreadsheet-id"),
            month=3,
        )
    except ValueError as exc:
        assert "No transactions were found for month 3" in str(exc)
    else:
        raise AssertionError("Expected empty pipeline result error")


def test_runlog_header_value_timestamp_is_not_treated_as_active_lock(monkeypatch, tmp_path):
    bank_file = tmp_path / "bank.xlsx"
    bank_file.write_text("x", encoding="utf-8")
    gateway = FakeSheetsGateway(lock_value="Timestamp")

    monkeypatch.setattr("ledgerflow.desktop_runner.build_sheets_gateway", lambda service_account_file=None: gateway)
    monkeypatch.setattr("ledgerflow.desktop_runner._build_openai_client", lambda cb: None)
    monkeypatch.setenv("LEDGERFLOW_YEAR_WORKSHEET_NAMES", '{"3":"ConfiguredMarch"}')
    monkeypatch.setattr(
        "ledgerflow.desktop_runner.extract_from_files_with_stats",
        lambda paths, config: ([], ExtractionStats()),
    )
    monkeypatch.setattr(
        "ledgerflow.desktop_runner.run_pipeline_from_config",
        lambda **kwargs: PipelineResult([], 0, 0, 0, 0, 0, 0, 0, [], False),
    )

    try:
        run_desktop_flow(
            bank_files=[bank_file],
            seed_file=Path("seed-spreadsheet-id"),
            yearly_workbook=Path("year-spreadsheet-id"),
            month=3,
        )
    except ValueError as exc:
        assert "No transactions were found for month 3" in str(exc)
    else:
        raise AssertionError("Expected empty pipeline result error")


def test_successful_runlog_append_does_not_interfere_with_lock_detection(monkeypatch, tmp_path):
    bank_file = tmp_path / "bank.xlsx"
    bank_file.write_text("x", encoding="utf-8")
    gateway = FakeSheetsGateway(lock_value="Timestamp")

    monkeypatch.setattr("ledgerflow.desktop_runner.build_sheets_gateway", lambda service_account_file=None: gateway)
    monkeypatch.setattr("ledgerflow.desktop_runner._build_openai_client", lambda cb: None)
    monkeypatch.setenv("LEDGERFLOW_YEAR_WORKSHEET_NAMES", '{"3":"ConfiguredMarch"}')
    monkeypatch.setattr(
        "ledgerflow.desktop_runner.extract_from_files_with_stats",
        lambda paths, config: ([sentinel.transaction], ExtractionStats(extracted_count=1)),
    )
    monkeypatch.setattr(
        "ledgerflow.desktop_runner.run_pipeline_from_config",
        lambda **kwargs: PipelineResult([sentinel.transaction], 1, 1, 1, 0, 0, 0, 1, [], False),
    )
    monkeypatch.setattr(
        "ledgerflow.desktop_runner.write_yearly_sheet",
        lambda **kwargs: WriteResult("not-applicable", "sheet-url", 1, 0),
    )
    monkeypatch.setattr("ledgerflow.desktop_runner.extract_gpt_transactions", lambda txs: [])

    summary = run_desktop_flow(
        bank_files=[bank_file],
        seed_file=Path("seed-spreadsheet-id"),
        yearly_workbook=Path("year-spreadsheet-id"),
        month=3,
    )

    assert summary.filtered_count == 1
    assert gateway.run_log_rows[0][0] == "Timestamp"
    assert gateway.lock_value == ""


def test_read_lock_value_ignores_non_lock_text():
    gateway = FakeSheetsGateway(lock_value="Timestamp")
    assert _read_lock_value(gateway, "year-spreadsheet-id", "RunLog") == ""


def test_read_lock_value_accepts_valid_lock_token():
    gateway = FakeSheetsGateway(lock_value="LOCK|host-a|123|2026-04-07 11:00:00")
    assert _read_lock_value(gateway, "year-spreadsheet-id", "RunLog") == "LOCK|host-a|123|2026-04-07 11:00:00"
