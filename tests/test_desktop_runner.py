from __future__ import annotations

from pathlib import Path
from unittest.mock import sentinel

from ledgerflow.desktop_runner import run_desktop_flow
from ledgerflow.extraction import ExtractionStats
from ledgerflow.pipeline import PipelineResult
from ledgerflow.year_writer import WriteResult


def test_run_desktop_flow_orchestrates_backend_modules(monkeypatch, tmp_path):
    bank_file = tmp_path / "bank.xlsx"
    seed_file = tmp_path / "seed.xlsx"
    yearly_workbook = tmp_path / "yearly.xlsx"
    for path in [bank_file, seed_file, yearly_workbook]:
        path.write_text("x", encoding="utf-8")

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

    def fake_write(*, transactions, source_workbook, output_path, month, config, backup_dir):
        captured["write_transactions"] = transactions
        captured["write_source_workbook"] = source_workbook
        captured["write_output_path"] = output_path
        captured["write_month"] = month
        captured["write_config"] = config
        captured["write_backup_dir"] = backup_dir
        return WriteResult(
            backup_path=backup_dir / "backup.xlsx",
            output_path=output_path,
            transactions_written=7,
            skipped_count=0,
        )

    monkeypatch.setattr("ledgerflow.desktop_runner.extract_from_files_with_stats", fake_extract)
    monkeypatch.setattr("ledgerflow.desktop_runner.run_pipeline_from_config", fake_pipeline)
    monkeypatch.setattr("ledgerflow.desktop_runner.write_yearly_workbook", fake_write)
    monkeypatch.setattr("ledgerflow.desktop_runner._build_openai_client", lambda cb: sentinel.gpt_client)
    monkeypatch.setattr("ledgerflow.desktop_runner.extract_gpt_transactions", lambda txs: [])

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
    assert captured["write_source_workbook"] == yearly_workbook
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
    assert summary.final_output_workbook_path == captured["write_output_path"]
    assert any("Extracting transactions" in status for status in statuses)
    assert any("Running categorization pipeline" in status for status in statuses)
    assert any("Writing yearly workbook output" in status for status in statuses)


def test_run_desktop_flow_rejects_empty_month_result(monkeypatch, tmp_path):
    bank_file = tmp_path / "bank.xlsx"
    seed_file = tmp_path / "seed.xlsx"
    yearly_workbook = tmp_path / "yearly.xlsx"
    for path in [bank_file, seed_file, yearly_workbook]:
        path.write_text("x", encoding="utf-8")

    monkeypatch.setattr(
        "ledgerflow.desktop_runner.extract_from_files_with_stats",
        lambda paths, config: ([], ExtractionStats(skipped_invalid_date=2)),
    )
    monkeypatch.setattr("ledgerflow.desktop_runner._build_openai_client", lambda cb: None)

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
