"""Tests for ledgerflow.csv_writer."""
from __future__ import annotations

import csv
from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest

from ledgerflow.csv_writer import (
    CATEGORIZED_COLUMNS,
    RAW_COLUMNS,
    write_categorized_csv,
    write_raw_csv,
)
from ledgerflow.models import Category, ClassificationSource, Transaction


# ------------------------------------------------------------------ #
# Helpers                                                              #
# ------------------------------------------------------------------ #


def tx(
    merchant: str = "Amazon",
    normalized: str = "amazon",
    d: str = "2024-03-15",
    amount: str = "99.99",
    source: str = "bank.xlsx",
    sheet: str = "Sheet1",
    category: Category | None = None,
    classification_source: ClassificationSource | None = None,
) -> Transaction:
    return Transaction(
        date=date.fromisoformat(d),
        merchant=merchant,
        normalized_merchant=normalized,
        amount=Decimal(amount),
        source_file=Path(source),
        sheet_name=sheet,
        category=category,
        classification_source=classification_source,
    )


def read_csv(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    """Return (headers, rows) from *path*."""
    with path.open(newline="", encoding="utf-8-sig") as fh:
        reader = csv.DictReader(fh)
        headers = list(reader.fieldnames or [])
        rows = list(reader)
    return headers, rows


# ------------------------------------------------------------------ #
# RAW_COLUMNS / CATEGORIZED_COLUMNS constants                         #
# ------------------------------------------------------------------ #


class TestColumnConstants:
    def test_raw_columns_are_stable(self):
        assert RAW_COLUMNS == (
            "date", "merchant", "normalized_merchant",
            "amount", "source_file", "sheet_name",
        )

    def test_categorized_columns_are_stable(self):
        assert CATEGORIZED_COLUMNS == (
            "date", "merchant", "normalized_merchant",
            "amount", "category", "classification_source", "source_file", "sheet_name",
        )

    def test_categorized_extends_raw_with_category(self):
        # Category column is inserted between amount and source_file
        raw = list(RAW_COLUMNS)
        cat = list(CATEGORIZED_COLUMNS)
        assert "category" in cat
        assert "category" not in raw


# ------------------------------------------------------------------ #
# write_raw_csv                                                        #
# ------------------------------------------------------------------ #


class TestWriteRawCsv:
    def test_creates_file(self, tmp_path):
        path = tmp_path / "raw.csv"
        write_raw_csv([tx()], path)
        assert path.exists()

    def test_header_row_matches_columns(self, tmp_path):
        path = tmp_path / "raw.csv"
        write_raw_csv([tx()], path)
        headers, _ = read_csv(path)
        assert headers == list(RAW_COLUMNS)

    def test_single_row_values(self, tmp_path):
        path = tmp_path / "raw.csv"
        write_raw_csv([tx()], path)
        _, rows = read_csv(path)
        assert len(rows) == 1
        r = rows[0]
        assert r["date"] == "2024-03-15"
        assert r["merchant"] == "Amazon"
        assert r["normalized_merchant"] == "amazon"
        assert r["amount"] == "99.99"
        assert r["sheet_name"] == "Sheet1"

    def test_source_file_included(self, tmp_path):
        path = tmp_path / "raw.csv"
        write_raw_csv([tx(source="data/bank_a.xlsx")], path)
        _, rows = read_csv(path)
        assert "bank_a.xlsx" in rows[0]["source_file"]

    def test_multiple_rows_preserved(self, tmp_path):
        path = tmp_path / "raw.csv"
        txs = [tx(merchant="Amazon"), tx(merchant="Spotify")]
        write_raw_csv(txs, path)
        _, rows = read_csv(path)
        assert len(rows) == 2
        assert rows[0]["merchant"] == "Amazon"
        assert rows[1]["merchant"] == "Spotify"

    def test_row_order_preserved(self, tmp_path):
        path = tmp_path / "raw.csv"
        txs = [tx(d="2024-03-15"), tx(d="2024-03-01"), tx(d="2024-03-28")]
        write_raw_csv(txs, path)
        _, rows = read_csv(path)
        dates = [r["date"] for r in rows]
        assert dates == ["2024-03-15", "2024-03-01", "2024-03-28"]

    def test_empty_list_writes_header_only(self, tmp_path):
        path = tmp_path / "raw.csv"
        write_raw_csv([], path)
        headers, rows = read_csv(path)
        assert headers == list(RAW_COLUMNS)
        assert rows == []

    def test_creates_parent_directories(self, tmp_path):
        path = tmp_path / "sub" / "dir" / "raw.csv"
        write_raw_csv([tx()], path)
        assert path.exists()

    def test_overwrites_existing_file(self, tmp_path):
        path = tmp_path / "raw.csv"
        write_raw_csv([tx(merchant="Old")], path)
        write_raw_csv([tx(merchant="New")], path)
        _, rows = read_csv(path)
        assert len(rows) == 1
        assert rows[0]["merchant"] == "New"

    def test_amount_is_exact_decimal_string(self, tmp_path):
        path = tmp_path / "raw.csv"
        write_raw_csv([tx(amount="1234.56")], path)
        _, rows = read_csv(path)
        assert rows[0]["amount"] == "1234.56"

    def test_hebrew_merchant_encoded_correctly(self, tmp_path):
        path = tmp_path / "raw.csv"
        write_raw_csv([tx(merchant="סופר פארם", normalized="סופר פארם")], path)
        _, rows = read_csv(path)
        assert rows[0]["merchant"] == "סופר פארם"

    def test_date_format_is_iso(self, tmp_path):
        path = tmp_path / "raw.csv"
        write_raw_csv([tx(d="2024-12-31")], path)
        _, rows = read_csv(path)
        assert rows[0]["date"] == "2024-12-31"

    def test_no_category_column(self, tmp_path):
        path = tmp_path / "raw.csv"
        write_raw_csv([tx()], path)
        headers, _ = read_csv(path)
        assert "category" not in headers


# ------------------------------------------------------------------ #
# write_categorized_csv                                                #
# ------------------------------------------------------------------ #


class TestWriteCategorizedCsv:
    def test_creates_file(self, tmp_path):
        path = tmp_path / "cat.csv"
        write_categorized_csv([tx(category=Category.BUSINESS)], path)
        assert path.exists()

    def test_header_row_matches_columns(self, tmp_path):
        path = tmp_path / "cat.csv"
        write_categorized_csv([tx(category=Category.BUSINESS)], path)
        headers, _ = read_csv(path)
        assert headers == list(CATEGORIZED_COLUMNS)

    def test_category_value_written(self, tmp_path):
        path = tmp_path / "cat.csv"
        write_categorized_csv(
            [tx(category=Category.LEISURE, classification_source=ClassificationSource.CACHE)],
            path,
        )
        _, rows = read_csv(path)
        assert rows[0]["category"] == "Leisure"
        assert rows[0]["classification_source"] == "cache"

    def test_all_category_values_written_correctly(self, tmp_path):
        path = tmp_path / "cat.csv"
        cats = [
            Category.OTHER, Category.BUSINESS, Category.PHARM,
            Category.INSURANCE, Category.TRANSPORTATION,
            Category.LEISURE, Category.HOME, Category.MORTGAGE,
        ]
        txs = [tx(merchant=c.value, category=c) for c in cats]
        write_categorized_csv(txs, path)
        _, rows = read_csv(path)
        written_cats = [r["category"] for r in rows]
        assert written_cats == [c.value for c in cats]

    def test_none_category_written_as_empty_string(self, tmp_path):
        path = tmp_path / "cat.csv"
        write_categorized_csv([tx(category=None)], path)
        _, rows = read_csv(path)
        assert rows[0]["category"] == ""
        assert rows[0]["classification_source"] == ""

    def test_all_base_columns_present(self, tmp_path):
        path = tmp_path / "cat.csv"
        write_categorized_csv([tx(category=Category.OTHER)], path)
        headers, _ = read_csv(path)
        for col in RAW_COLUMNS:
            assert col in headers

    def test_empty_list_writes_header_only(self, tmp_path):
        path = tmp_path / "cat.csv"
        write_categorized_csv([], path)
        headers, rows = read_csv(path)
        assert headers == list(CATEGORIZED_COLUMNS)
        assert rows == []

    def test_creates_parent_directories(self, tmp_path):
        path = tmp_path / "out" / "final.csv"
        write_categorized_csv([tx(category=Category.OTHER)], path)
        assert path.exists()

    def test_hebrew_category_not_used_but_hebrew_merchant_works(self, tmp_path):
        path = tmp_path / "cat.csv"
        write_categorized_csv(
            [tx(merchant="סופר פארם", normalized="סופר פארם", category=Category.PHARM)],
            path,
        )
        _, rows = read_csv(path)
        assert rows[0]["merchant"] == "סופר פארם"
        assert rows[0]["category"] == "Pharm"
