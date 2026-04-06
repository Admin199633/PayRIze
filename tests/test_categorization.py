"""
Tests for ledgerflow.categorization.

Strategy:
- load_seed_mapping tests build real .xlsx files in tmp_path (openpyxl).
- All categorize_transactions / cache / utility tests are pure in-memory.
- Fixtures are kept minimal and explicit.
"""
from __future__ import annotations

import json
import pytest
from datetime import date
from decimal import Decimal
from pathlib import Path

import openpyxl

from ledgerflow.categorization import (
    CategorizationResult,
    _is_fuel_station_merchant,
    categorize_transactions,
    load_cache,
    load_seed_mapping,
    merge_mappings,
    save_cache,
    uncategorized_merchants,
)
from ledgerflow.config import Config
from ledgerflow.models import Category, ClassificationSource, Transaction


# ------------------------------------------------------------------ #
# Shared helpers                                                       #
# ------------------------------------------------------------------ #


def make_tx(
    merchant: str,
    normalized: str | None = None,
    d: str = "2024-03-15",
    amount: str = "100.00",
    source: str = "bank.xlsx",
    sheet: str = "Sheet1",
) -> Transaction:
    return Transaction(
        date=date.fromisoformat(d),
        merchant=merchant,
        normalized_merchant=normalized if normalized is not None else merchant.lower().strip(),
        amount=Decimal(amount),
        source_file=Path(source),
        sheet_name=sheet,
    )


def write_seed(
    tmp_path: Path,
    rows: list[dict],
    headers: tuple[str, str, str] = ("MerchantKey", "MerchantOriginal", "Category"),
    name: str = "seed.xlsx",
) -> Path:
    """
    Write a seed Excel file.

    *rows* is a list of dicts whose keys are the *canonical* column names
    (MerchantKey, MerchantOriginal, Category).  *headers* controls what is
    actually written to the header row so non-standard casing/spacing can be
    tested.  Row values are always fetched by the canonical keys.
    """
    _CANONICAL = ("MerchantKey", "MerchantOriginal", "Category")
    path = tmp_path / name
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(list(headers))
    for row in rows:
        ws.append([row.get(canonical) for canonical in _CANONICAL])
    wb.save(str(path))
    return path


# ------------------------------------------------------------------ #
# load_seed_mapping                                                    #
# ------------------------------------------------------------------ #


class TestLoadSeedMapping:
    def test_basic_load_returns_mapping(self, tmp_path):
        path = write_seed(tmp_path, [
            {"MerchantKey": "amazon", "MerchantOriginal": "Amazon.com", "Category": "Business"},
        ])
        mapping = load_seed_mapping(path)
        assert mapping["amazon"] == Category.BUSINESS

    def test_merchantoriginal_also_indexed(self, tmp_path):
        path = write_seed(tmp_path, [
            {"MerchantKey": "amzn", "MerchantOriginal": "Amazon.com", "Category": "Business"},
        ])
        mapping = load_seed_mapping(path)
        # "Amazon.com" → lowercase → dot stripped → "amazoncom" (letters run together)
        assert "amazoncom" in mapping
        assert mapping["amazoncom"] == Category.BUSINESS

    def test_merchantkey_also_indexed(self, tmp_path):
        path = write_seed(tmp_path, [
            {"MerchantKey": "spotify ab", "MerchantOriginal": "Spotify", "Category": "Leisure"},
        ])
        mapping = load_seed_mapping(path)
        assert "spotify ab" in mapping
        assert mapping["spotify ab"] == Category.LEISURE

    def test_both_keys_point_to_same_category(self, tmp_path):
        path = write_seed(tmp_path, [
            {"MerchantKey": "wmt", "MerchantOriginal": "Walmart Supercenter", "Category": "Home"},
        ])
        mapping = load_seed_mapping(path)
        assert mapping.get("wmt") == Category.HOME
        assert mapping.get("walmart supercenter") == Category.HOME

    def test_normalization_applied_to_keys(self, tmp_path):
        # Raw key "AMAZON.COM" normalizes to "amazoncom"
        path = write_seed(tmp_path, [
            {"MerchantKey": "AMAZON.COM", "MerchantOriginal": "AMAZON.COM", "Category": "Business"},
        ])
        mapping = load_seed_mapping(path)
        assert "amazoncom" in mapping

    def test_case_insensitive_header_matching(self, tmp_path):
        # Headers in all-caps should still be recognised
        path = write_seed(
            tmp_path,
            [{"MerchantKey": "shop", "MerchantOriginal": "Shop Inc", "Category": "Business"}],
            headers=("MERCHANTKEY", "MERCHANTORIGINAL", "CATEGORY"),
        )
        mapping = load_seed_mapping(path)
        assert "shop" in mapping

    def test_headers_with_surrounding_whitespace(self, tmp_path):
        path = write_seed(
            tmp_path,
            [{"MerchantKey": "shop", "MerchantOriginal": "Shop", "Category": "Business"}],
            headers=("  MerchantKey  ", "  MerchantOriginal  ", "  Category  "),
        )
        mapping = load_seed_mapping(path)
        assert "shop" in mapping

    def test_invalid_category_row_skipped(self, tmp_path):
        path = write_seed(tmp_path, [
            {"MerchantKey": "good", "MerchantOriginal": "Good Shop", "Category": "Business"},
            {"MerchantKey": "bad", "MerchantOriginal": "Bad Shop", "Category": "NotACategory"},
        ])
        mapping = load_seed_mapping(path)
        assert "good" in mapping
        assert "bad" not in mapping

    def test_none_category_row_skipped(self, tmp_path):
        path = write_seed(tmp_path, [
            {"MerchantKey": "shop", "MerchantOriginal": "Shop", "Category": None},
        ])
        mapping = load_seed_mapping(path)
        assert "shop" not in mapping

    def test_none_merchantkey_skips_key_entry(self, tmp_path):
        path = write_seed(tmp_path, [
            {"MerchantKey": None, "MerchantOriginal": "Amazon", "Category": "Business"},
        ])
        mapping = load_seed_mapping(path)
        # MerchantOriginal still indexed; MerchantKey is absent
        assert "amazon" in mapping

    def test_none_merchantoriginal_skips_original_entry(self, tmp_path):
        path = write_seed(tmp_path, [
            {"MerchantKey": "amzn", "MerchantOriginal": None, "Category": "Business"},
        ])
        mapping = load_seed_mapping(path)
        assert "amzn" in mapping

    def test_all_valid_categories_accepted(self, tmp_path):
        rows = [
            {"MerchantKey": k, "MerchantOriginal": k, "Category": cat.value}
            for k, cat in [
                ("shop_other", Category.OTHER),
                ("shop_biz", Category.BUSINESS),
                ("shop_pharm", Category.PHARM),
                ("shop_ins", Category.INSURANCE),
                ("shop_trans", Category.TRANSPORTATION),
                ("shop_leisure", Category.LEISURE),
                ("shop_home", Category.HOME),
                ("shop_mortgage", Category.MORTGAGE),
            ]
        ]
        path = write_seed(tmp_path, rows)
        mapping = load_seed_mapping(path)
        assert mapping["shop_other"] == Category.OTHER
        assert mapping["shop_mortgage"] == Category.MORTGAGE

    def test_missing_merchantkey_column_raises(self, tmp_path):
        path = write_seed(
            tmp_path,
            [{"MerchantKey": "x", "MerchantOriginal": "X", "Category": "Business"}],
            headers=("WRONG", "MerchantOriginal", "Category"),
        )
        with pytest.raises(ValueError, match="MerchantKey"):
            load_seed_mapping(path)

    def test_missing_merchantoriginal_column_raises(self, tmp_path):
        path = write_seed(
            tmp_path,
            [{"MerchantKey": "x", "MerchantOriginal": "X", "Category": "Business"}],
            headers=("MerchantKey", "WRONG", "Category"),
        )
        with pytest.raises(ValueError, match="MerchantOriginal"):
            load_seed_mapping(path)

    def test_missing_category_column_raises(self, tmp_path):
        path = write_seed(
            tmp_path,
            [{"MerchantKey": "x", "MerchantOriginal": "X", "Category": "Business"}],
            headers=("MerchantKey", "MerchantOriginal", "WRONG"),
        )
        with pytest.raises(ValueError, match="Category"):
            load_seed_mapping(path)

    def test_all_columns_missing_lists_all_in_error(self, tmp_path):
        path = write_seed(
            tmp_path,
            [{"MerchantKey": "x", "MerchantOriginal": "X", "Category": "Business"}],
            headers=("A", "B", "C"),
        )
        with pytest.raises(ValueError) as exc_info:
            load_seed_mapping(path)
        msg = str(exc_info.value)
        assert "MerchantKey" in msg
        assert "MerchantOriginal" in msg
        assert "Category" in msg

    def test_empty_workbook_raises(self, tmp_path):
        path = tmp_path / "empty.xlsx"
        wb = openpyxl.Workbook()
        wb.active.append([])
        wb.save(str(path))
        with pytest.raises(ValueError):
            load_seed_mapping(path)

    def test_returns_empty_dict_when_no_data_rows(self, tmp_path):
        path = write_seed(tmp_path, [])  # header row only, no data
        mapping = load_seed_mapping(path)
        assert mapping == {}

    def test_file_not_found_propagates(self, tmp_path):
        with pytest.raises(Exception):
            load_seed_mapping(tmp_path / "nonexistent.xlsx")


# ------------------------------------------------------------------ #
# load_cache / save_cache                                              #
# ------------------------------------------------------------------ #


class TestCache:
    def test_load_nonexistent_returns_empty(self, tmp_path):
        assert load_cache(tmp_path / "nope.json") == {}

    def test_round_trip(self, tmp_path):
        path = tmp_path / "cache.json"
        mapping = {"amazon": Category.BUSINESS, "spotify": Category.LEISURE}
        save_cache(mapping, path)
        loaded = load_cache(path)
        assert loaded == mapping

    def test_invalid_category_string_dropped(self, tmp_path):
        path = tmp_path / "cache.json"
        path.write_text(
            json.dumps({"amazon": "Business", "bad": "NotARealCategory"}),
            encoding="utf-8",
        )
        loaded = load_cache(path)
        assert "amazon" in loaded
        assert "bad" not in loaded

    def test_save_creates_parent_directories(self, tmp_path):
        path = tmp_path / "sub" / "dir" / "cache.json"
        save_cache({"shop": Category.HOME}, path)
        assert path.exists()

    def test_cache_json_is_human_readable(self, tmp_path):
        path = tmp_path / "cache.json"
        save_cache({"amazon": Category.BUSINESS}, path)
        text = path.read_text(encoding="utf-8")
        assert "Business" in text  # enum value, not enum name
        assert "amazon" in text


# ------------------------------------------------------------------ #
# categorize_transactions                                              #
# ------------------------------------------------------------------ #


class TestCategorizeTransactions:
    # ---- basic resolution ----

    def test_seed_resolves_transaction(self):
        txs = [make_tx("Amazon", normalized="amazon")]
        seed = {"amazon": Category.BUSINESS}
        result = categorize_transactions(txs, seed, {}, Config())
        assert result.transactions[0].category == Category.BUSINESS
        assert result.transactions[0].classification_source == ClassificationSource.SEED
        assert result.seed_count == 1
        assert result.cache_count == 0
        assert result.fallback_count == 0

    def test_cache_resolves_when_not_in_seed(self):
        txs = [make_tx("Spotify", normalized="spotify")]
        cache = {"spotify": Category.LEISURE}
        result = categorize_transactions(txs, {}, cache, Config())
        assert result.transactions[0].category == Category.LEISURE
        assert result.transactions[0].classification_source == ClassificationSource.CACHE
        assert result.cache_count == 1
        assert result.seed_count == 0

    def test_fallback_to_other_when_no_mapping(self):
        txs = [make_tx("Unknown Shop", normalized="unknown shop")]
        result = categorize_transactions(txs, {}, {}, Config())
        assert result.transactions[0].category == Category.OTHER
        assert result.transactions[0].classification_source == ClassificationSource.FALLBACK
        assert result.fallback_count == 1

    def test_gpt_merchants_marked_as_gpt(self):
        txs = [make_tx("Unknown Shop", normalized="unknown shop")]
        cache = {"unknown shop": Category.LEISURE}
        result = categorize_transactions(
            txs, {}, cache, Config(), gpt_merchants={"unknown shop"}
        )
        assert result.transactions[0].category == Category.LEISURE
        assert result.transactions[0].classification_source == ClassificationSource.GPT

    # ---- precedence ----

    def test_seed_takes_precedence_over_cache(self):
        txs = [make_tx("Amazon", normalized="amazon")]
        seed = {"amazon": Category.BUSINESS}
        cache = {"amazon": Category.OTHER}  # lower trust; should be ignored
        result = categorize_transactions(txs, seed, cache, Config())
        assert result.transactions[0].category == Category.BUSINESS
        assert result.seed_count == 1
        assert result.cache_count == 0

    def test_cache_used_only_when_seed_has_no_entry(self):
        txs = [
            make_tx("Amazon", normalized="amazon"),
            make_tx("Spotify", normalized="spotify"),
        ]
        seed = {"amazon": Category.BUSINESS}
        cache = {"spotify": Category.LEISURE}
        result = categorize_transactions(txs, seed, cache, Config())
        cats = {tx.normalized_merchant: tx.category for tx in result.transactions}
        assert cats["amazon"] == Category.BUSINESS
        assert cats["spotify"] == Category.LEISURE
        assert result.seed_count == 1
        assert result.cache_count == 1

    # ---- invalid / disallowed categories ----

    def test_disallowed_seed_category_falls_through_to_cache(self):
        # Mortgage in seed but not allowed → should use cache value instead.
        cfg = Config()
        cfg.include_mortgage = False
        txs = [make_tx("Bank", normalized="bank")]
        seed = {"bank": Category.MORTGAGE}
        cache = {"bank": Category.HOME}
        result = categorize_transactions(txs, seed, cache, cfg)
        assert result.transactions[0].category == Category.HOME
        assert result.cache_count == 1
        assert result.seed_count == 0

    def test_disallowed_seed_and_cache_falls_to_other(self):
        cfg = Config()
        cfg.include_mortgage = False
        txs = [make_tx("Bank", normalized="bank")]
        seed = {"bank": Category.MORTGAGE}
        cache = {"bank": Category.MORTGAGE}
        result = categorize_transactions(txs, seed, cache, cfg)
        assert result.transactions[0].category == Category.OTHER
        assert result.fallback_count == 1
        assert "bank" in result.unresolved_merchants

    def test_invalid_cache_category_not_in_enum_is_already_filtered_by_load_cache(
        self, tmp_path
    ):
        # load_cache drops non-enum strings, so by the time we call
        # categorize_transactions the cache only has valid Category values.
        path = tmp_path / "cache.json"
        path.write_text(
            json.dumps({"amazon": "NotACategory"}), encoding="utf-8"
        )
        cache = load_cache(path)
        assert "amazon" not in cache  # dropped at load time

        txs = [make_tx("Amazon", normalized="amazon")]
        result = categorize_transactions(txs, {}, cache, Config())
        assert result.transactions[0].category == Category.OTHER

    # ---- Mortgage gating ----

    def test_mortgage_allowed_when_config_enables_it(self):
        cfg = Config()
        cfg.include_mortgage = True
        cfg.categories.append(Category.MORTGAGE)
        txs = [make_tx("Mortgage Co", normalized="mortgage co")]
        seed = {"mortgage co": Category.MORTGAGE}
        result = categorize_transactions(txs, seed, {}, cfg)
        assert result.transactions[0].category == Category.MORTGAGE

    def test_mortgage_rejected_when_config_disables_it(self):
        cfg = Config()
        cfg.include_mortgage = False
        txs = [make_tx("Mortgage Co", normalized="mortgage co")]
        seed = {"mortgage co": Category.MORTGAGE}
        result = categorize_transactions(txs, seed, {}, cfg)
        assert result.transactions[0].category == Category.OTHER
        assert result.fallback_count == 1

    # ---- unresolved_merchants ----

    def test_unresolved_merchants_listed_when_no_mapping(self):
        txs = [
            make_tx("Unknown A", normalized="unknown a"),
            make_tx("Unknown B", normalized="unknown b"),
        ]
        result = categorize_transactions(txs, {}, {}, Config())
        assert "unknown a" in result.unresolved_merchants
        assert "unknown b" in result.unresolved_merchants

    def test_explicitly_mapped_to_other_is_not_unresolved(self):
        # A merchant mapped to Other in the seed is "resolved", not unresolved.
        txs = [make_tx("Misc", normalized="misc")]
        seed = {"misc": Category.OTHER}
        result = categorize_transactions(txs, seed, {}, Config())
        assert result.transactions[0].category == Category.OTHER
        assert "misc" not in result.unresolved_merchants
        assert result.seed_count == 1
        assert result.fallback_count == 0

    def test_unresolved_merchants_are_sorted(self):
        txs = [
            make_tx("Z Shop", normalized="z shop"),
            make_tx("A Shop", normalized="a shop"),
        ]
        result = categorize_transactions(txs, {}, {}, Config())
        assert result.unresolved_merchants == sorted(result.unresolved_merchants)

    def test_unresolved_merchants_deduplicated(self):
        # Same merchant appears in two transactions — should appear once.
        txs = [
            make_tx("Amazon", normalized="amazon"),
            make_tx("Amazon", normalized="amazon"),
        ]
        result = categorize_transactions(txs, {}, {}, Config())
        assert result.unresolved_merchants.count("amazon") == 1

    # ---- counts ----

    def test_counts_sum_to_total_transactions(self):
        txs = [
            make_tx("A", normalized="a"),  # seed
            make_tx("B", normalized="b"),  # cache
            make_tx("C", normalized="c"),  # fallback
        ]
        seed = {"a": Category.BUSINESS}
        cache = {"b": Category.LEISURE}
        result = categorize_transactions(txs, seed, cache, Config())
        total = result.seed_count + result.cache_count + result.fallback_count
        assert total == len(txs)

    def test_counts_for_empty_input(self):
        result = categorize_transactions([], {}, {}, Config())
        assert result.seed_count == 0
        assert result.cache_count == 0
        assert result.fallback_count == 0
        assert result.transactions == []
        assert result.unresolved_merchants == []

    # ---- immutability ----

    def test_input_transactions_not_mutated(self):
        txs = [make_tx("Amazon", normalized="amazon")]
        assert txs[0].category is None

        seed = {"amazon": Category.BUSINESS}
        result = categorize_transactions(txs, seed, {}, Config())

        # Original list is untouched.
        assert txs[0].category is None
        # Result has a new object with category set.
        assert result.transactions[0].category == Category.BUSINESS
        assert result.transactions[0] is not txs[0]

    def test_result_contains_all_original_fields(self):
        tx = make_tx("Amazon", normalized="amazon", d="2024-06-15", amount="199.99")
        seed = {"amazon": Category.BUSINESS}
        result = categorize_transactions([tx], seed, {}, Config())
        out = result.transactions[0]
        assert out.date == tx.date
        assert out.merchant == tx.merchant
        assert out.normalized_merchant == tx.normalized_merchant
        assert out.amount == tx.amount
        assert out.source_file == tx.source_file
        assert out.sheet_name == tx.sheet_name

    # ---- normalization consistency ----

    def test_normalized_key_used_for_lookup(self):
        # Transaction's normalized_merchant must match the seed key.
        # If normalization was applied correctly during extraction, "AMAZON.COM"
        # → normalized "amazoncom"; seed key should be "amazoncom".
        txs = [make_tx("AMAZON.COM", normalized="amazoncom")]
        seed = {"amazoncom": Category.BUSINESS}
        result = categorize_transactions(txs, seed, {}, Config())
        assert result.transactions[0].category == Category.BUSINESS

    def test_raw_merchant_is_not_used_for_lookup(self):
        # Only normalized_merchant is used; raw merchant differences are ignored.
        txs = [make_tx("AMAZON.COM !!!!", normalized="amazon com")]
        seed = {"amazon com": Category.BUSINESS}
        result = categorize_transactions(txs, seed, {}, Config())
        assert result.transactions[0].category == Category.BUSINESS


# ------------------------------------------------------------------ #
# merge_mappings                                                       #
# ------------------------------------------------------------------ #


class TestMergeMappings:
    def test_single_mapping_passthrough(self):
        m = {"amazon": Category.BUSINESS}
        assert merge_mappings(m) == m

    def test_later_overrides_earlier(self):
        m1 = {"amazon": Category.OTHER}
        m2 = {"amazon": Category.BUSINESS}
        result = merge_mappings(m1, m2)
        assert result["amazon"] == Category.BUSINESS

    def test_non_overlapping_keys_all_present(self):
        m1 = {"amazon": Category.BUSINESS}
        m2 = {"spotify": Category.LEISURE}
        result = merge_mappings(m1, m2)
        assert result["amazon"] == Category.BUSINESS
        assert result["spotify"] == Category.LEISURE

    def test_empty_inputs(self):
        assert merge_mappings({}, {}) == {}

    def test_does_not_mutate_inputs(self):
        m1 = {"amazon": Category.OTHER}
        m2 = {"amazon": Category.BUSINESS}
        merge_mappings(m1, m2)
        assert m1["amazon"] == Category.OTHER  # unchanged


# ------------------------------------------------------------------ #
# uncategorized_merchants                                              #
# ------------------------------------------------------------------ #


class TestUncategorizedMerchants:
    def test_all_missing(self):
        txs = [
            make_tx("Amazon", normalized="amazon"),
            make_tx("Spotify", normalized="spotify"),
        ]
        result = uncategorized_merchants(txs, {})
        assert result == ["amazon", "spotify"]

    def test_none_missing(self):
        txs = [make_tx("Amazon", normalized="amazon")]
        mapping = {"amazon": Category.BUSINESS}
        assert uncategorized_merchants(txs, mapping) == []

    def test_partial_missing(self):
        txs = [
            make_tx("Amazon", normalized="amazon"),
            make_tx("Unknown", normalized="unknown"),
        ]
        mapping = {"amazon": Category.BUSINESS}
        result = uncategorized_merchants(txs, mapping)
        assert result == ["unknown"]

    def test_returns_sorted_list(self):
        txs = [
            make_tx("Z", normalized="z"),
            make_tx("A", normalized="a"),
        ]
        result = uncategorized_merchants(txs, {})
        assert result == ["a", "z"]

    def test_duplicates_deduplicated(self):
        txs = [
            make_tx("Amazon", normalized="amazon"),
            make_tx("Amazon", normalized="amazon"),
        ]
        result = uncategorized_merchants(txs, {})
        assert result == ["amazon"]

    def test_empty_transactions(self):
        assert uncategorized_merchants([], {}) == []


# ------------------------------------------------------------------ #
# _is_fuel_station_merchant                                            #
# ------------------------------------------------------------------ #


class TestIsFuelStationMerchant:
    # ── positive matches ──────────────────────────────────────────────

    def test_paz_yellow_variant(self):
        assert _is_fuel_station_merchant("פז yellow צבר") is True

    def test_paz_alone(self):
        assert _is_fuel_station_merchant("פז דלק") is True

    def test_yellow_alone(self):
        assert _is_fuel_station_merchant("yellow") is True

    def test_delek_retail(self):
        assert _is_fuel_station_merchant("דלק קמעונאות דרור") is True

    def test_delek_menta(self):
        assert _is_fuel_station_merchant("דלק מנטה ארז מודיעין") is True

    def test_sonol(self):
        assert _is_fuel_station_merchant("סונול שלי") is True

    def test_dor_alon(self):
        assert _is_fuel_station_merchant("דור אלון") is True

    def test_alonit(self):
        assert _is_fuel_station_merchant("אלונית") is True

    def test_super_alonit(self):
        assert _is_fuel_station_merchant("סופר אלונית אלון דב") is True

    # ── negative matches — unrelated merchants ────────────────────────

    def test_regular_supermarket(self):
        assert _is_fuel_station_merchant("סופרמרקט שופרסל") is False

    def test_amazon(self):
        assert _is_fuel_station_merchant("amazon") is False

    def test_empty_string(self):
        assert _is_fuel_station_merchant("") is False

    # ── transit exceptions override fuel keywords ─────────────────────

    def test_train_not_fuel(self):
        assert _is_fuel_station_merchant("רכבת ישראל") is False

    def test_bus_not_fuel(self):
        assert _is_fuel_station_merchant("אוטובוס 480") is False

    def test_rav_kav_not_fuel(self):
        assert _is_fuel_station_merchant("רב קו") is False

    def test_parking_not_fuel(self):
        assert _is_fuel_station_merchant("חניה מרכזית") is False

    def test_license_renewal_not_fuel(self):
        assert _is_fuel_station_merchant("חידוש רישיון") is False

    def test_toll_road_not_fuel(self):
        assert _is_fuel_station_merchant("כביש 6 אגרה") is False

    def test_pango_not_fuel(self):
        assert _is_fuel_station_merchant("פנגו חניה") is False


# ------------------------------------------------------------------ #
# Fuel convenience rule — integration via categorize_transactions      #
# ------------------------------------------------------------------ #


def _fuel_tx(merchant: str, normalized: str, amount: str) -> Transaction:
    """Create a transaction pre-seeded as Transportation via seed mapping."""
    return make_tx(merchant, normalized=normalized, amount=amount)


class TestFuelConvenienceRule:
    """
    The fuel rule fires ONLY when:
      category == Transportation AND amount < 50 AND fuel-station merchant.
    Result: category → Leisure, source → RULE.
    """

    def _seed_as_transport(self, normalized: str) -> dict:
        return {normalized: Category.TRANSPORTATION}

    # ── threshold boundary ────────────────────────────────────────────

    def test_amount_49_90_becomes_leisure(self):
        norm = "פז yellow צבר"
        txs = [make_tx("פז YELLOW צבר", normalized=norm, amount="49.90")]
        result = categorize_transactions(txs, self._seed_as_transport(norm), {}, Config())
        assert result.transactions[0].category == Category.LEISURE
        assert result.transactions[0].classification_source == ClassificationSource.RULE

    def test_amount_50_00_stays_transportation(self):
        norm = "פז yellow צבר"
        txs = [make_tx("פז YELLOW צבר", normalized=norm, amount="50.00")]
        result = categorize_transactions(txs, self._seed_as_transport(norm), {}, Config())
        assert result.transactions[0].category == Category.TRANSPORTATION
        assert result.transactions[0].classification_source == ClassificationSource.SEED

    def test_amount_50_01_stays_transportation(self):
        norm = "דלק קמעונאות דרור"
        txs = [make_tx("דלק קמעונאות דרור", normalized=norm, amount="50.01")]
        result = categorize_transactions(txs, self._seed_as_transport(norm), {}, Config())
        assert result.transactions[0].category == Category.TRANSPORTATION

    def test_amount_0_01_becomes_leisure(self):
        norm = "סונול שלי"
        txs = [make_tx("סונול שלי", normalized=norm, amount="0.01")]
        result = categorize_transactions(txs, self._seed_as_transport(norm), {}, Config())
        assert result.transactions[0].category == Category.LEISURE

    # ── transit exceptions stay Transportation ────────────────────────

    def test_train_amount_20_stays_transportation(self):
        norm = "רכבת ישראל"
        txs = [make_tx("רכבת ישראל", normalized=norm, amount="20.00")]
        result = categorize_transactions(txs, self._seed_as_transport(norm), {}, Config())
        assert result.transactions[0].category == Category.TRANSPORTATION
        assert result.transactions[0].classification_source == ClassificationSource.SEED

    def test_license_renewal_amount_30_stays_transportation(self):
        norm = "חידוש רישיון רכב"
        txs = [make_tx("חידוש רישיון רכב", normalized=norm, amount="30.00")]
        result = categorize_transactions(txs, self._seed_as_transport(norm), {}, Config())
        assert result.transactions[0].category == Category.TRANSPORTATION

    def test_parking_amount_15_stays_transportation(self):
        norm = "חניה מרכזית תל אביב"
        txs = [make_tx("חניה מרכזית", normalized=norm, amount="15.00")]
        result = categorize_transactions(txs, self._seed_as_transport(norm), {}, Config())
        assert result.transactions[0].category == Category.TRANSPORTATION

    def test_toll_road_stays_transportation(self):
        norm = "כביש 6 אגרה"
        txs = [make_tx("כביש 6", normalized=norm, amount="18.00")]
        result = categorize_transactions(txs, self._seed_as_transport(norm), {}, Config())
        assert result.transactions[0].category == Category.TRANSPORTATION

    # ── non-Transportation categories are unaffected ──────────────────

    def test_other_category_fuel_merchant_unchanged(self):
        norm = "פז yellow"
        txs = [make_tx("פז YELLOW", normalized=norm, amount="30.00")]
        result = categorize_transactions(txs, {"פז yellow": Category.OTHER}, {}, Config())
        assert result.transactions[0].category == Category.OTHER

    def test_leisure_category_fuel_merchant_unchanged(self):
        norm = "סונול"
        txs = [make_tx("סונול", normalized=norm, amount="25.00")]
        result = categorize_transactions(txs, {norm: Category.LEISURE}, {}, Config())
        assert result.transactions[0].category == Category.LEISURE

    # ── merchant variants covered ─────────────────────────────────────

    def test_paz_yellow_tzabar_variant(self):
        norm = "פז yellow צבר"
        txs = [make_tx("פז YELLOW צבר", normalized=norm, amount="35.00")]
        result = categorize_transactions(txs, self._seed_as_transport(norm), {}, Config())
        assert result.transactions[0].category == Category.LEISURE

    def test_delek_retail_dror_variant(self):
        norm = "דלק קמעונאות דרור"
        txs = [make_tx("דלק קמעונאות דרור", normalized=norm, amount="22.00")]
        result = categorize_transactions(txs, self._seed_as_transport(norm), {}, Config())
        assert result.transactions[0].category == Category.LEISURE

    def test_sonol_hamaccabim_variant(self):
        norm = "סונול המכבים"
        txs = [make_tx("סונול המכבים", normalized=norm, amount="19.50")]
        result = categorize_transactions(txs, self._seed_as_transport(norm), {}, Config())
        assert result.transactions[0].category == Category.LEISURE

    def test_alonit_variant(self):
        norm = "סופר אלונית אלון דב"
        txs = [make_tx("סופר אלונית אלון דב", normalized=norm, amount="44.00")]
        result = categorize_transactions(txs, self._seed_as_transport(norm), {}, Config())
        assert result.transactions[0].category == Category.LEISURE

    # ── rule fires regardless of original classification source ───────

    def test_rule_overrides_seed(self):
        norm = "דלק קמעונאות רמה"
        txs = [make_tx("דלק קמעונאות רמה", normalized=norm, amount="25.00")]
        seed = {norm: Category.TRANSPORTATION}
        result = categorize_transactions(txs, seed, {}, Config())
        assert result.transactions[0].category == Category.LEISURE
        assert result.transactions[0].classification_source == ClassificationSource.RULE

    def test_rule_overrides_cache(self):
        norm = "אלונית עטרות"
        txs = [make_tx("אלונית עטרות", normalized=norm, amount="18.00")]
        cache = {norm: Category.TRANSPORTATION}
        result = categorize_transactions(txs, {}, cache, Config())
        assert result.transactions[0].category == Category.LEISURE
        assert result.transactions[0].classification_source == ClassificationSource.RULE

    # ── non-fuel merchant with same amount is unaffected ─────────────

    def test_non_fuel_transport_merchant_stays_transport(self):
        norm = "גוש דן אוטובוסים"
        txs = [make_tx("גוש דן", normalized=norm, amount="10.00")]
        result = categorize_transactions(txs, self._seed_as_transport(norm), {}, Config())
        # "אוטובוס" pattern matches → transit exception → stays Transportation
        assert result.transactions[0].category == Category.TRANSPORTATION

    def test_unrelated_merchant_low_amount_unchanged(self):
        norm = "בית קפה ירושלים"
        txs = [make_tx("בית קפה", normalized=norm, amount="25.00")]
        result = categorize_transactions(txs, {norm: Category.LEISURE}, {}, Config())
        assert result.transactions[0].category == Category.LEISURE
