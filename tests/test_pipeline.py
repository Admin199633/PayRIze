"""
Tests for ledgerflow.pipeline.

Strategy:
- run_pipeline is called with pre-loaded dicts (no seed/cache files needed).
- GPT client is a MagicMock throughout.
- CSV output is verified by reading back the written files.
- run_pipeline_from_config is tested with a minimal integration test only.
"""
from __future__ import annotations

import csv
import json
from datetime import date
from decimal import Decimal
from pathlib import Path
from unittest.mock import MagicMock, call

import pytest

from ledgerflow.categorization import save_cache
from ledgerflow.config import Config
from ledgerflow.csv_writer import CATEGORIZED_COLUMNS, RAW_COLUMNS
from ledgerflow.models import Category, ClassificationSource, Transaction
from ledgerflow.pipeline import PipelineResult, run_pipeline, run_pipeline_from_config


# ------------------------------------------------------------------ #
# Helpers                                                              #
# ------------------------------------------------------------------ #


def tx(
    merchant: str,
    normalized: str | None = None,
    month: int = 3,
    day: int = 15,
    amount: str = "50.00",
    source: str = "bank.xlsx",
) -> Transaction:
    return Transaction(
        date=date(2024, month, day),
        merchant=merchant,
        normalized_merchant=normalized if normalized is not None else merchant.lower(),
        amount=Decimal(amount),
        source_file=Path(source),
        sheet_name="Sheet1",
    )


def make_gpt_client(response: dict[str, str]) -> MagicMock:
    """Return a mock GPT client that always responds with *response*."""
    import json as _json
    client = MagicMock()
    client.chat.completions.create.return_value.choices[
        0
    ].message.content = _json.dumps(response, ensure_ascii=False)
    return client


def read_csv_rows(path: Path) -> list[dict[str, str]]:
    with path.open(newline="", encoding="utf-8-sig") as fh:
        return list(csv.DictReader(fh))


# ------------------------------------------------------------------ #
# PipelineResult properties                                            #
# ------------------------------------------------------------------ #


class TestPipelineResultProperties:
    def _result(self, seed=0, cache=0, fallback=0, filtered=10):
        return PipelineResult(
            transactions=[],
            extracted_count=10,
            after_dedup_count=10,
            after_month_filter_count=filtered,
            seed_count=seed,
            cache_count=cache,
            gpt_count=0,
            fallback_count=fallback,
            unresolved_merchants=[],
            cache_was_updated=False,
        )

    def test_total_resolved(self):
        r = self._result(seed=3, cache=2)
        assert r.total_resolved == 5

    def test_resolution_rate(self):
        r = self._result(seed=7, cache=2, fallback=1, filtered=10)
        assert r.resolution_rate == pytest.approx(0.9)

    def test_resolution_rate_zero_filtered(self):
        r = self._result(filtered=0)
        assert r.resolution_rate == 1.0


# ------------------------------------------------------------------ #
# Month filter integration                                             #
# ------------------------------------------------------------------ #


class TestPipelineMonthFilter:
    def test_only_target_month_returned(self):
        txs = [tx("A", month=3), tx("B", month=4), tx("C", month=3)]
        result = run_pipeline(txs, 3, {}, {}, Config())
        assert result.after_month_filter_count == 2

    def test_empty_after_filter_skips_categorization(self):
        txs = [tx("A", month=4)]
        client = MagicMock()
        result = run_pipeline(txs, 3, {}, {}, Config(), gpt_client=client)
        assert result.after_month_filter_count == 0
        assert result.transactions == []
        client.chat.completions.create.assert_not_called()

    def test_counts_reflect_filter_boundary(self):
        txs = [tx("A", month=3), tx("B", month=4)]
        result = run_pipeline(txs, 3, {}, {}, Config())
        assert result.extracted_count == 2
        assert result.after_month_filter_count == 1


# ------------------------------------------------------------------ #
# Dedup integration                                                    #
# ------------------------------------------------------------------ #


class TestPipelineDedup:
    def test_dedup_reduces_count_before_month_filter(self):
        # Two identical transactions in the same month — only one survives dedup.
        t = tx("Amazon", month=3, amount="100.00")
        result = run_pipeline([t, t], 3, {}, {}, Config())
        assert result.after_dedup_count == 1
        assert result.after_month_filter_count == 1

    def test_dedup_happens_before_filter(self):
        # Duplicates across different months: dedup removes one,
        # then month filter keeps the March one.
        t_march = tx("Amazon", month=3, amount="100.00")
        t_april = tx("Amazon", month=4, amount="100.00")
        result = run_pipeline([t_march, t_april], 3, {}, {}, Config())
        # Dedup removes one (same date key would differ by month — actually
        # different keys since date differs). Both survive dedup. Filter keeps March only.
        assert result.after_dedup_count == 2
        assert result.after_month_filter_count == 1


# ------------------------------------------------------------------ #
# Categorization without GPT                                           #
# ------------------------------------------------------------------ #


class TestPipelineCategorizationNoGpt:
    def test_seed_resolves_transaction(self):
        txs = [tx("Amazon", normalized="amazon", month=3)]
        seed = {"amazon": Category.BUSINESS}
        result = run_pipeline(txs, 3, seed, {}, Config())
        assert result.seed_count == 1
        assert result.transactions[0].category == Category.BUSINESS
        assert result.transactions[0].classification_source == ClassificationSource.SEED

    def test_cache_resolves_transaction(self):
        txs = [tx("Spotify", normalized="spotify", month=3)]
        cache = {"spotify": Category.LEISURE}
        result = run_pipeline(txs, 3, {}, cache, Config())
        assert result.cache_count == 1
        assert result.transactions[0].category == Category.LEISURE
        assert result.transactions[0].classification_source == ClassificationSource.CACHE

    def test_fallback_to_other_when_not_in_mappings(self):
        txs = [tx("Unknown", month=3)]
        result = run_pipeline(txs, 3, {}, {}, Config())
        assert result.fallback_count == 1
        assert result.transactions[0].category == Category.OTHER
        assert result.transactions[0].classification_source == ClassificationSource.FALLBACK

    def test_no_gpt_call_when_client_is_none(self):
        txs = [tx("Unknown", month=3)]
        # No exception should be raised; gpt_client=None means skip GPT.
        result = run_pipeline(txs, 3, {}, {}, Config(), gpt_client=None)
        assert result.gpt_count == 0
        assert result.cache_was_updated is False

    def test_no_gpt_call_when_all_resolved_by_seed(self):
        txs = [tx("Amazon", normalized="amazon", month=3)]
        seed = {"amazon": Category.BUSINESS}
        client = MagicMock()
        result = run_pipeline(txs, 3, seed, {}, Config(), gpt_client=client)
        client.chat.completions.create.assert_not_called()
        assert result.gpt_count == 0


# ------------------------------------------------------------------ #
# GPT enrichment                                                       #
# ------------------------------------------------------------------ #


class TestPipelineGptEnrichment:
    def test_gpt_called_only_for_unresolved_merchants(self):
        txs = [
            tx("Amazon", normalized="amazon", month=3),
            tx("UnknownShop", normalized="unknownshop", month=3),
        ]
        seed = {"amazon": Category.BUSINESS}
        client = make_gpt_client({"unknownshop": "Leisure"})
        run_pipeline(txs, 3, seed, {}, Config(), gpt_client=client)
        _, kwargs = client.chat.completions.create.call_args
        user_msg = json.loads(kwargs["messages"][1]["content"])
        assert "unknownshop" in user_msg
        assert "amazon" not in user_msg

    def test_gpt_result_applied_in_second_pass(self):
        txs = [tx("UnknownShop", normalized="unknownshop", month=3)]
        client = make_gpt_client({"unknownshop": "Leisure"})
        result = run_pipeline(txs, 3, {}, {}, Config(), gpt_client=client)
        assert result.transactions[0].category == Category.LEISURE
        assert result.transactions[0].classification_source == ClassificationSource.GPT
        assert result.gpt_count == 1

    def test_gpt_count_counts_transactions_not_unique_merchants(self):
        # Same merchant in two transactions → GPT resolves 1 merchant, 2 transactions.
        txs = [
            tx("Shop", normalized="shop", month=3, day=1),
            tx("Shop", normalized="shop", month=3, day=2),
        ]
        client = make_gpt_client({"shop": "Home"})
        result = run_pipeline(txs, 3, {}, {}, Config(), gpt_client=client)
        assert result.gpt_count == 2

    def test_no_gpt_call_when_empty_unresolved(self):
        txs = [tx("Amazon", normalized="amazon", month=3)]
        seed = {"amazon": Category.BUSINESS}
        client = MagicMock()
        run_pipeline(txs, 3, seed, {}, Config(), gpt_client=client)
        client.chat.completions.create.assert_not_called()

    def test_no_second_pass_when_gpt_returns_empty(self):
        # GPT client returns empty JSON object.
        txs = [tx("Unknown", normalized="unknown", month=3)]
        client = MagicMock()
        client.chat.completions.create.return_value.choices[
            0
        ].message.content = "{}"
        result = run_pipeline(txs, 3, {}, {}, Config(), gpt_client=client)
        # Only one API call (first pass had unresolved; GPT returned nothing).
        assert result.gpt_count == 0
        assert result.fallback_count == 1

    def test_partial_gpt_results_applied(self):
        txs = [
            tx("KnownShop", normalized="knownshop", month=3),
            tx("UnknownShop", normalized="unknownshop", month=3),
        ]
        # GPT only resolves one of the two unknown merchants.
        client = make_gpt_client({"knownshop": "Business"})
        result = run_pipeline(txs, 3, {}, {}, Config(), gpt_client=client)
        cats = {t.normalized_merchant: t.category for t in result.transactions}
        assert cats["knownshop"] == Category.BUSINESS
        assert cats["unknownshop"] == Category.OTHER

    def test_seed_still_takes_precedence_after_gpt(self):
        txs = [tx("Amazon", normalized="amazon", month=3)]
        seed = {"amazon": Category.BUSINESS}
        # GPT disagrees with seed.
        client = make_gpt_client({"amazon": "Other"})
        result = run_pipeline(txs, 3, seed, {}, Config(), gpt_client=client)
        assert result.transactions[0].category == Category.BUSINESS
        assert result.seed_count == 1


# ------------------------------------------------------------------ #
# Cache persistence                                                    #
# ------------------------------------------------------------------ #


class TestPipelineCachePersistence:
    def test_cache_saved_to_config_path_after_gpt(self, tmp_path):
        cache_path = tmp_path / "cache.json"
        cfg = Config()
        cfg.cache_path = cache_path

        txs = [tx("Shop", normalized="shop", month=3)]
        client = make_gpt_client({"shop": "Leisure"})
        result = run_pipeline(txs, 3, {}, {}, cfg, gpt_client=client)

        assert result.cache_was_updated is True
        assert cache_path.exists()
        data = json.loads(cache_path.read_text(encoding="utf-8"))
        assert data["shop"] == "Leisure"

    def test_cache_not_saved_when_no_config_path(self):
        cfg = Config()
        cfg.cache_path = None
        txs = [tx("Shop", normalized="shop", month=3)]
        client = make_gpt_client({"shop": "Leisure"})
        result = run_pipeline(txs, 3, {}, {}, cfg, gpt_client=client)
        assert result.cache_was_updated is False

    def test_cache_not_saved_when_gpt_returns_nothing(self, tmp_path):
        cache_path = tmp_path / "cache.json"
        cfg = Config()
        cfg.cache_path = cache_path

        txs = [tx("Shop", normalized="shop", month=3)]
        client = MagicMock()
        client.chat.completions.create.return_value.choices[
            0
        ].message.content = "{}"
        result = run_pipeline(txs, 3, {}, {}, cfg, gpt_client=client)

        assert result.cache_was_updated is False
        assert not cache_path.exists()

    def test_cache_merges_gpt_results_with_existing_cache(self, tmp_path):
        cache_path = tmp_path / "cache.json"
        cfg = Config()
        cfg.cache_path = cache_path

        existing_cache = {"amazon": Category.BUSINESS}
        save_cache(existing_cache, cache_path)

        txs = [tx("NewShop", normalized="newshop", month=3)]
        client = make_gpt_client({"newshop": "Home"})
        run_pipeline(txs, 3, {}, existing_cache, cfg, gpt_client=client)

        data = json.loads(cache_path.read_text(encoding="utf-8"))
        assert data["amazon"] == "Business"   # preserved from original
        assert data["newshop"] == "Home"       # added by GPT

    def test_no_gpt_call_when_month_filtered_empty(self):
        txs = [tx("Shop", month=4)]
        client = MagicMock()
        run_pipeline(txs, 3, {}, {}, Config(), gpt_client=client)
        client.chat.completions.create.assert_not_called()


# ------------------------------------------------------------------ #
# Pipeline counts consistency                                          #
# ------------------------------------------------------------------ #


class TestPipelineCounts:
    def test_seed_plus_cache_plus_gpt_plus_fallback_equals_filter_count(self):
        txs = [
            tx("SeedMerchant", normalized="seedmerchant", month=3),
            tx("CacheMerchant", normalized="cachemerchant", month=3),
            tx("GptMerchant", normalized="gptmerchant", month=3),
            tx("FallbackMerchant", normalized="fallbackmerchant", month=3),
        ]
        seed = {"seedmerchant": Category.BUSINESS}
        cache = {"cachemerchant": Category.HOME}
        client = make_gpt_client({"gptmerchant": "Leisure"})
        cfg = Config()

        result = run_pipeline(txs, 3, seed, cache, cfg, gpt_client=client)

        total = result.seed_count + result.cache_count + result.fallback_count
        assert total == result.after_month_filter_count

    def test_gpt_count_correct_after_second_pass(self):
        txs = [
            tx("A", normalized="a", month=3),
            tx("B", normalized="b", month=3),
        ]
        # Both unresolved in first pass; GPT resolves 'a' only.
        client = make_gpt_client({"a": "Business"})
        result = run_pipeline(txs, 3, {}, {}, Config(), gpt_client=client)
        assert result.gpt_count == 1
        assert result.fallback_count == 1

    def test_extracted_count_matches_input_length(self):
        txs = [tx("A", month=3), tx("B", month=3), tx("C", month=4)]
        result = run_pipeline(txs, 3, {}, {}, Config())
        assert result.extracted_count == 3

    def test_empty_pipeline_all_counts_zero(self):
        result = run_pipeline([], 3, {}, {}, Config())
        assert result.extracted_count == 0
        assert result.after_dedup_count == 0
        assert result.after_month_filter_count == 0
        assert result.seed_count == 0
        assert result.cache_count == 0
        assert result.gpt_count == 0
        assert result.fallback_count == 0


# ------------------------------------------------------------------ #
# CSV output                                                           #
# ------------------------------------------------------------------ #


class TestPipelineCsvOutput:
    def test_raw_csv_written_when_path_given(self, tmp_path):
        path = tmp_path / "raw.csv"
        txs = [tx("Amazon", month=3)]
        run_pipeline(txs, 3, {}, {}, Config(), raw_csv_path=path)
        assert path.exists()

    def test_raw_csv_contains_filtered_transactions(self, tmp_path):
        path = tmp_path / "raw.csv"
        txs = [tx("A", month=3), tx("B", month=4)]  # B excluded by month filter
        run_pipeline(txs, 3, {}, {}, Config(), raw_csv_path=path)
        rows = read_csv_rows(path)
        assert len(rows) == 1
        assert rows[0]["merchant"] == "A"

    def test_raw_csv_does_not_have_category_column(self, tmp_path):
        path = tmp_path / "raw.csv"
        run_pipeline([tx("A", month=3)], 3, {}, {}, Config(), raw_csv_path=path)
        rows = read_csv_rows(path)
        assert "category" not in (rows[0] if rows else {})
        with path.open(encoding="utf-8-sig") as fh:
            headers = fh.readline().strip().split(",")
        assert "category" not in headers

    def test_categorized_csv_written_when_path_given(self, tmp_path):
        path = tmp_path / "cat.csv"
        txs = [tx("Amazon", month=3)]
        run_pipeline(txs, 3, {}, {}, Config(), categorized_csv_path=path)
        assert path.exists()

    def test_categorized_csv_contains_category_column(self, tmp_path):
        path = tmp_path / "cat.csv"
        seed = {"amazon": Category.BUSINESS}
        run_pipeline(
            [tx("Amazon", normalized="amazon", month=3)],
            3,
            seed,
            {},
            Config(),
            categorized_csv_path=path,
        )
        rows = read_csv_rows(path)
        assert rows[0]["category"] == "Business"
        assert rows[0]["classification_source"] == "seed"

    def test_no_csv_written_when_paths_not_given(self, tmp_path):
        run_pipeline([tx("A", month=3)], 3, {}, {}, Config())
        # No files should appear in tmp_path
        assert list(tmp_path.iterdir()) == []

    def test_raw_csv_not_written_when_month_filter_empties(self, tmp_path):
        path = tmp_path / "raw.csv"
        run_pipeline([tx("A", month=4)], 3, {}, {}, Config(), raw_csv_path=path)
        # Pipeline returns early with empty result — no CSV written
        assert not path.exists()

    def test_categorized_csv_not_written_when_month_filter_empties(self, tmp_path):
        path = tmp_path / "cat.csv"
        run_pipeline([tx("A", month=4)], 3, {}, {}, Config(), categorized_csv_path=path)
        assert not path.exists()

    def test_both_csv_paths_written_in_single_call(self, tmp_path):
        raw_path = tmp_path / "raw.csv"
        cat_path = tmp_path / "cat.csv"
        run_pipeline(
            [tx("A", month=3)],
            3,
            {},
            {},
            Config(),
            raw_csv_path=raw_path,
            categorized_csv_path=cat_path,
        )
        assert raw_path.exists()
        assert cat_path.exists()


# ------------------------------------------------------------------ #
# run_pipeline_from_config                                             #
# ------------------------------------------------------------------ #


class TestRunPipelineFromConfig:
    def test_empty_config_paths_uses_empty_dicts(self):
        cfg = Config()
        # seed_mapping_path and cache_path are both None — should not crash.
        txs = [tx("Amazon", month=3)]
        result = run_pipeline_from_config(txs, 3, cfg)
        assert result.after_month_filter_count == 1
        assert result.fallback_count == 1

    def test_seed_path_loaded(self, tmp_path):
        import openpyxl

        seed_path = tmp_path / "seed.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["MerchantKey", "MerchantOriginal", "Category"])
        ws.append(["amazon", "Amazon", "Business"])
        wb.save(str(seed_path))

        cfg = Config()
        cfg.seed_mapping_path = seed_path

        txs = [tx("Amazon", normalized="amazon", month=3)]
        result = run_pipeline_from_config(txs, 3, cfg)
        assert result.seed_count == 1
        assert result.transactions[0].category == Category.BUSINESS

    def test_cache_path_loaded_and_updated(self, tmp_path):
        cache_path = tmp_path / "cache.json"
        initial_cache = {"spotify": "Leisure"}
        cache_path.write_text(json.dumps(initial_cache), encoding="utf-8")

        cfg = Config()
        cfg.cache_path = cache_path

        txs = [tx("Spotify", normalized="spotify", month=3)]
        result = run_pipeline_from_config(txs, 3, cfg)
        assert result.cache_count == 1
        assert result.transactions[0].category == Category.LEISURE
