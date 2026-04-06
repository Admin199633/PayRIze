"""Tests for ledgerflow.seed_updater."""
from __future__ import annotations

from datetime import date
from decimal import Decimal
from pathlib import Path

import openpyxl
import pytest

from ledgerflow.models import Category, ClassificationSource, Transaction
from ledgerflow.seed_updater import (
    NUMERIC_CATEGORY_MAP,
    category_number,
    extract_gpt_transactions,
    update_seed_file,
)


# ------------------------------------------------------------------ #
# Helpers                                                              #
# ------------------------------------------------------------------ #


def _tx(
    merchant: str = "Shop",
    source: ClassificationSource = ClassificationSource.GPT,
    category: Category = Category.OTHER,
) -> Transaction:
    return Transaction(
        date=date(2024, 3, 15),
        merchant=merchant,
        normalized_merchant=merchant.lower(),
        amount=Decimal("10.00"),
        source_file=Path("bank.xlsx"),
        sheet_name="Sheet1",
        category=category,
        classification_source=source,
    )


def _make_seed(tmp_path: Path, rows: list[tuple]) -> Path:
    """Create a seed workbook with given data rows (after a standard header)."""
    path = tmp_path / "seed.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["MerchantKey", "MerchantOriginal", "Category"])
    for row in rows:
        ws.append(list(row))
    wb.save(str(path))
    return path


def _read_seed_rows(path: Path) -> list[tuple]:
    """Return all data rows (after the first header row) as plain tuples."""
    wb = openpyxl.load_workbook(str(path))
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    return [tuple(r) for r in rows[1:]]  # skip header


# ------------------------------------------------------------------ #
# NUMERIC_CATEGORY_MAP                                                 #
# ------------------------------------------------------------------ #


class TestNumericCategoryMap:
    def test_all_eight_keys_present(self):
        assert set(NUMERIC_CATEGORY_MAP.keys()) == {1, 2, 3, 4, 5, 6, 7, 8}

    def test_key_1_is_home(self):
        assert NUMERIC_CATEGORY_MAP[1] == Category.HOME

    def test_key_2_is_leisure(self):
        assert NUMERIC_CATEGORY_MAP[2] == Category.LEISURE

    def test_key_3_is_transportation(self):
        assert NUMERIC_CATEGORY_MAP[3] == Category.TRANSPORTATION

    def test_key_4_is_insurance(self):
        assert NUMERIC_CATEGORY_MAP[4] == Category.INSURANCE

    def test_key_5_is_pharm(self):
        assert NUMERIC_CATEGORY_MAP[5] == Category.PHARM

    def test_key_6_is_business(self):
        assert NUMERIC_CATEGORY_MAP[6] == Category.BUSINESS

    def test_key_7_is_mortgage(self):
        assert NUMERIC_CATEGORY_MAP[7] == Category.MORTGAGE

    def test_key_8_is_other(self):
        assert NUMERIC_CATEGORY_MAP[8] == Category.OTHER

    def test_all_values_are_category_instances(self):
        for num, cat in NUMERIC_CATEGORY_MAP.items():
            assert isinstance(cat, Category), f"Key {num} maps to non-Category {cat!r}"

    def test_no_duplicate_categories(self):
        values = list(NUMERIC_CATEGORY_MAP.values())
        assert len(values) == len(set(values))


class TestCategoryNumber:
    def test_home_returns_1(self):
        assert category_number(Category.HOME) == 1

    def test_leisure_returns_2(self):
        assert category_number(Category.LEISURE) == 2

    def test_other_returns_8(self):
        assert category_number(Category.OTHER) == 8

    def test_round_trip_for_all_mapped_categories(self):
        for num, cat in NUMERIC_CATEGORY_MAP.items():
            assert category_number(cat) == num


# ------------------------------------------------------------------ #
# extract_gpt_transactions                                             #
# ------------------------------------------------------------------ #


class TestExtractGptTransactions:
    def test_returns_only_gpt_transactions(self):
        txs = [
            _tx(merchant="A", source=ClassificationSource.GPT),
            _tx(merchant="B", source=ClassificationSource.SEED),
            _tx(merchant="C", source=ClassificationSource.GPT),
            _tx(merchant="D", source=ClassificationSource.CACHE),
            _tx(merchant="E", source=ClassificationSource.FALLBACK),
        ]
        result = extract_gpt_transactions(txs)
        assert len(result) == 2
        assert all(tx.classification_source == ClassificationSource.GPT for tx in result)
        merchants = [tx.merchant for tx in result]
        assert "A" in merchants
        assert "C" in merchants

    def test_empty_list_returns_empty(self):
        assert extract_gpt_transactions([]) == []

    def test_no_gpt_returns_empty(self):
        txs = [
            _tx(source=ClassificationSource.SEED),
            _tx(source=ClassificationSource.CACHE),
            _tx(source=ClassificationSource.FALLBACK),
        ]
        assert extract_gpt_transactions(txs) == []

    def test_all_gpt_returns_all(self):
        txs = [_tx(merchant=m, source=ClassificationSource.GPT) for m in "ABC"]
        assert len(extract_gpt_transactions(txs)) == 3

    def test_none_classification_source_excluded(self):
        tx = _tx()
        tx2 = Transaction(
            date=date(2024, 3, 15),
            merchant="NoSource",
            normalized_merchant="nosource",
            amount=Decimal("10.00"),
            source_file=Path("bank.xlsx"),
            sheet_name="Sheet1",
            category=Category.OTHER,
            classification_source=None,
        )
        assert extract_gpt_transactions([tx2]) == []

    def test_preserves_transaction_order(self):
        txs = [
            _tx(merchant="First", source=ClassificationSource.GPT),
            _tx(merchant="Second", source=ClassificationSource.SEED),
            _tx(merchant="Third", source=ClassificationSource.GPT),
        ]
        result = extract_gpt_transactions(txs)
        assert result[0].merchant == "First"
        assert result[1].merchant == "Third"


# ------------------------------------------------------------------ #
# update_seed_file — new file creation                                 #
# ------------------------------------------------------------------ #


class TestUpdateSeedFileCreate:
    def test_creates_file_when_not_exists(self, tmp_path):
        path = tmp_path / "new_seed.xlsx"
        update_seed_file(path, {"amazon": ("Amazon", Category.BUSINESS)})
        assert path.exists()

    def test_created_file_has_header_row(self, tmp_path):
        path = tmp_path / "seed.xlsx"
        update_seed_file(path, {"shop": ("Shop", Category.OTHER)})
        wb = openpyxl.load_workbook(str(path))
        ws = wb.active
        row1 = [str(c).strip().lower() if c else "" for c in next(ws.iter_rows(values_only=True))]
        assert "merchantkey" in row1
        assert "merchantoriginal" in row1
        assert "category" in row1

    def test_single_entry_written(self, tmp_path):
        path = tmp_path / "seed.xlsx"
        update_seed_file(path, {"amazon": ("Amazon", Category.BUSINESS)})
        rows = _read_seed_rows(path)
        assert len(rows) == 1
        assert rows[0][0] == "amazon"
        assert rows[0][1] == "Amazon"
        assert rows[0][2] == "Business"

    def test_multiple_entries_written(self, tmp_path):
        path = tmp_path / "seed.xlsx"
        reviewed = {
            "amazon": ("Amazon", Category.BUSINESS),
            "netflix": ("Netflix", Category.LEISURE),
        }
        update_seed_file(path, reviewed)
        rows = _read_seed_rows(path)
        keys = {r[0] for r in rows}
        assert "amazon" in keys
        assert "netflix" in keys

    def test_empty_reviewed_does_not_create_file(self, tmp_path):
        path = tmp_path / "seed.xlsx"
        update_seed_file(path, {})
        assert not path.exists()

    def test_creates_parent_directories(self, tmp_path):
        path = tmp_path / "sub" / "dir" / "seed.xlsx"
        update_seed_file(path, {"shop": ("Shop", Category.OTHER)})
        assert path.exists()


# ------------------------------------------------------------------ #
# update_seed_file — update existing file                              #
# ------------------------------------------------------------------ #


class TestUpdateSeedFileUpdate:
    def test_new_entry_appended_to_existing_file(self, tmp_path):
        path = _make_seed(tmp_path, [("amazon", "Amazon", "Business")])
        update_seed_file(path, {"netflix": ("Netflix", Category.LEISURE)})
        rows = _read_seed_rows(path)
        keys = {r[0] for r in rows}
        assert "amazon" in keys
        assert "netflix" in keys

    def test_existing_entry_category_updated(self, tmp_path):
        path = _make_seed(tmp_path, [("amazon", "Amazon", "Business")])
        update_seed_file(path, {"amazon": ("Amazon", Category.OTHER)})
        rows = _read_seed_rows(path)
        # Should still be exactly one row for amazon
        amazon_rows = [r for r in rows if r[0] == "amazon"]
        assert len(amazon_rows) == 1
        assert amazon_rows[0][2] == "Other"

    def test_duplicate_not_created_for_existing_key(self, tmp_path):
        path = _make_seed(tmp_path, [("amazon", "Amazon", "Business")])
        update_seed_file(path, {"amazon": ("Amazon", Category.LEISURE)})
        rows = _read_seed_rows(path)
        amazon_rows = [r for r in rows if r[0] == "amazon"]
        assert len(amazon_rows) == 1

    def test_multiple_updates_and_appends(self, tmp_path):
        path = _make_seed(tmp_path, [
            ("amazon", "Amazon", "Business"),
            ("netflix", "Netflix", "Leisure"),
        ])
        update_seed_file(path, {
            "amazon": ("Amazon", Category.OTHER),   # update
            "spotify": ("Spotify", Category.LEISURE),  # new
        })
        rows = _read_seed_rows(path)
        by_key = {r[0]: r for r in rows if r[0]}
        assert by_key["amazon"][2] == "Other"
        assert by_key["netflix"][2] == "Leisure"  # unchanged
        assert by_key["spotify"][2] == "Leisure"  # new

    def test_empty_reviewed_leaves_file_unchanged(self, tmp_path):
        path = _make_seed(tmp_path, [("amazon", "Amazon", "Business")])
        original_rows = _read_seed_rows(path)
        update_seed_file(path, {})
        assert _read_seed_rows(path) == original_rows

    def test_first_occurrence_of_duplicate_key_is_updated(self, tmp_path):
        """Seed file with two rows sharing the same key — first is updated."""
        path = _make_seed(tmp_path, [
            ("amazon", "Amazon", "Business"),
            ("amazon", "AMAZON", "Other"),  # duplicate
        ])
        update_seed_file(path, {"amazon": ("Amazon Prime", Category.LEISURE)})
        rows = _read_seed_rows(path)
        # First occurrence updated, second left alone
        assert rows[0][2] == "Leisure"
        assert rows[1][2] == "Other"

    def test_written_category_value_matches_enum_value(self, tmp_path):
        path = tmp_path / "seed.xlsx"
        update_seed_file(path, {"pharm": ("Pharm", Category.PHARM)})
        rows = _read_seed_rows(path)
        assert rows[0][2] == Category.PHARM.value  # "Pharm"

    def test_all_categories_can_be_written(self, tmp_path):
        path = tmp_path / "seed.xlsx"
        reviewed = {cat.value.lower(): (cat.value, cat) for cat in Category}
        update_seed_file(path, reviewed)
        rows = _read_seed_rows(path)
        written_cats = {r[2] for r in rows if r[2]}
        for cat in Category:
            assert cat.value in written_cats
