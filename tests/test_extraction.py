"""
Tests for ledgerflow.extraction.

Strategy:
- Helper functions (detect_table_header, is_stop_row, etc.) are tested with
  plain Python lists — no workbook files needed, fast, and completely isolated.
- _extract_from_sheet is tested with plain Python lists too (it is pure).
- extract_from_file / extract_from_files get integration tests that build real
  .xlsx files in tmp_path using openpyxl directly.
"""
from __future__ import annotations

import pytest
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import openpyxl

from ledgerflow.config import Config
from ledgerflow.extraction import (
    ExtractionStats,
    UnsupportedFileTypeError,
    _extract_from_sheet,
    _extract_from_sheet_with_stats,
    detect_table_header,
    extract_from_file,
    extract_from_file_with_stats,
    extract_from_files,
    extract_from_files_with_stats,
    is_all_empty_row,
    is_stop_row,
    parse_transaction_row,
)


# ------------------------------------------------------------------ #
# Shared fixtures / helpers                                            #
# ------------------------------------------------------------------ #

@pytest.fixture
def cfg() -> Config:
    return Config()


# Stable source metadata used throughout unit tests.
_SRC = Path("test.xlsx")
_SHEET = "Sheet1"


def write_workbook(
    tmp_path: Path,
    sheets: dict[str, list[list]],
    name: str = "test.xlsx",
) -> Path:
    """
    Create a real .xlsx workbook at *tmp_path/name* and return its path.
    Each key in *sheets* becomes a sheet; each inner list becomes a row.
    """
    path = tmp_path / name
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # drop the default blank sheet
    for sheet_name, rows in sheets.items():
        ws = wb.create_sheet(title=sheet_name)
        for row in rows:
            ws.append(row)
    wb.save(str(path))
    return path


# ------------------------------------------------------------------ #
# detect_table_header                                                  #
# ------------------------------------------------------------------ #

class TestDetectTableHeader:
    def test_standard_english_headers(self, cfg):
        row = ["Purchase Date", "Merchant", "Amount"]
        assert detect_table_header(row, cfg) == (0, 1, 2)

    def test_alternative_english_aliases(self, cfg):
        row = ["Transaction Date", "Business Name", "Charge"]
        assert detect_table_header(row, cfg) == (0, 1, 2)

    def test_charge_date_alias(self, cfg):
        row = ["Charge Date", "Business Name", "Charge"]
        assert detect_table_header(row, cfg) == (0, 1, 2)

    def test_prefix_header_aliases_match(self, cfg):
        row = ["Date Posted", "Merchant Name", "Amount Total"]
        assert detect_table_header(row, cfg) == (0, 1, 2)

    def test_hebrew_headers(self, cfg):
        row = ["תאריך עסקה", "שם בית עסק", "סכום חיוב"]
        assert detect_table_header(row, cfg) == (0, 1, 2)

    def test_case_insensitive_matching(self, cfg):
        row = ["PURCHASE DATE", "MERCHANT", "AMOUNT"]
        assert detect_table_header(row, cfg) == (0, 1, 2)

    def test_whitespace_around_header_stripped(self, cfg):
        row = ["  Purchase Date  ", " Merchant ", " Amount "]
        assert detect_table_header(row, cfg) == (0, 1, 2)

    def test_non_standard_column_order(self, cfg):
        # merchant=0, amount=1, date=2
        row = ["Merchant", "Amount", "Purchase Date"]
        assert detect_table_header(row, cfg) == (2, 0, 1)

    def test_charge_amount_column_preferred_over_transaction_amount(self, cfg):
        row = [
            "תאריך עסקה",
            "שם בית עסק",
            "סכום עסקה",
            "מטבע עסקה",
            "סכום חיוב",
            "מטבע חיוב",
        ]
        assert detect_table_header(row, cfg) == (0, 1, 4)

    def test_charge_amount_header_variant_is_still_preferred(self, cfg):
        row = [
            "תאריך עסקה",
            "שם בית עסק",
            "סכום עסקה",
            "מטבע עסקה",
            "סכום חיוב בש\"ח",
            "מטבע חיוב",
        ]
        assert detect_table_header(row, cfg) == (0, 1, 4)

    def test_extra_columns_on_both_sides(self, cfg):
        row = ["Reference", "Purchase Date", "Merchant", "Amount", "Notes"]
        assert detect_table_header(row, cfg) == (1, 2, 3)

    def test_none_cells_are_skipped(self, cfg):
        row = [None, "Purchase Date", None, "Merchant", None, "Amount"]
        assert detect_table_header(row, cfg) == (1, 3, 5)

    def test_missing_date_column_returns_none(self, cfg):
        assert detect_table_header(["Merchant", "Amount"], cfg) is None

    def test_missing_merchant_column_returns_none(self, cfg):
        assert detect_table_header(["Purchase Date", "Amount"], cfg) is None

    def test_missing_amount_column_returns_none(self, cfg):
        assert detect_table_header(["Purchase Date", "Merchant"], cfg) is None

    def test_empty_row_returns_none(self, cfg):
        assert detect_table_header([], cfg) is None

    def test_all_none_row_returns_none(self, cfg):
        assert detect_table_header([None, None, None], cfg) is None

    def test_data_row_not_misidentified_as_header(self, cfg):
        row = ["15/03/2024", "Amazon", "99.99"]
        assert detect_table_header(row, cfg) is None


# ------------------------------------------------------------------ #
# is_stop_row                                                          #
# ------------------------------------------------------------------ #

class TestIsStopRow:
    def test_total_keyword_stops(self, cfg):
        assert is_stop_row(["Total", None, None], cfg) is True

    def test_grand_total_stops(self, cfg):
        assert is_stop_row([None, "Grand Total", 500], cfg) is True

    def test_subtotal_stops(self, cfg):
        assert is_stop_row(["Subtotal", None, "100"], cfg) is True

    def test_hebrew_total_stops(self, cfg):
        assert is_stop_row(['סה"כ', None, None], cfg) is True

    def test_case_insensitive(self, cfg):
        assert is_stop_row(["TOTAL", None, None], cfg) is True

    def test_token_as_substring(self, cfg):
        assert is_stop_row(["Grand Total Charges", None, None], cfg) is True

    def test_normal_data_row_is_not_stop(self, cfg):
        assert is_stop_row(["2024-03-15", "Amazon", "99.99"], cfg) is False

    def test_all_none_is_not_stop(self, cfg):
        assert is_stop_row([None, None, None], cfg) is False

    def test_empty_strings_are_not_stop(self, cfg):
        assert is_stop_row(["", "  ", ""], cfg) is False

    def test_custom_stop_token_respected(self):
        cfg = Config()
        cfg.stop_tokens = ["end of statement"]
        assert is_stop_row(["End of Statement", None, None], cfg) is True
        assert is_stop_row(["Total", None, None], cfg) is False

    def test_stop_token_in_any_column(self, cfg):
        # token anywhere in the row — not just the first cell
        assert is_stop_row(["2024-03-15", "Total charges", "500"], cfg) is True


# ------------------------------------------------------------------ #
# is_all_empty_row                                                     #
# ------------------------------------------------------------------ #

class TestIsAllEmptyRow:
    def test_all_none(self):
        assert is_all_empty_row([None, None, None]) is True

    def test_all_empty_strings(self):
        assert is_all_empty_row(["", "  ", "\t"]) is True

    def test_mixed_none_and_blank(self):
        assert is_all_empty_row([None, "", None]) is True

    def test_empty_list(self):
        assert is_all_empty_row([]) is True

    def test_row_with_string_value(self):
        assert is_all_empty_row([None, "Amazon", None]) is False

    def test_row_with_zero_is_not_empty(self):
        # 0 is a valid cell value, not empty
        assert is_all_empty_row([None, 0, None]) is False

    def test_row_with_false_is_not_empty(self):
        assert is_all_empty_row([None, False, None]) is False


# ------------------------------------------------------------------ #
# parse_transaction_row                                                #
# ------------------------------------------------------------------ #

class TestParseTransactionRow:
    def _p(self, cells, dc=0, mc=1, ac=2):
        return parse_transaction_row(cells, dc, mc, ac, _SRC, _SHEET)

    def test_valid_row_returns_transaction(self):
        tx = self._p(["15/03/2024", "Amazon", "99.99"])
        assert tx is not None
        assert tx.date == date(2024, 3, 15)
        assert tx.merchant == "Amazon"
        assert tx.normalized_merchant == "amazon"
        assert tx.amount == Decimal("99.99")
        assert tx.source_file == _SRC
        assert tx.sheet_name == _SHEET
        assert tx.category is None

    def test_datetime_object_date_column(self):
        tx = self._p([datetime(2024, 3, 15), "Walmart", "50.00"])
        assert tx is not None
        assert tx.date == date(2024, 3, 15)

    def test_invalid_date_returns_none(self):
        assert self._p(["not-a-date", "Amazon", "99.99"]) is None

    def test_none_date_returns_none(self):
        assert self._p([None, "Amazon", "99.99"]) is None

    def test_invalid_amount_returns_none(self):
        assert self._p(["15/03/2024", "Amazon", "N/A"]) is None

    def test_none_amount_returns_none(self):
        assert self._p(["15/03/2024", "Amazon", None]) is None

    def test_empty_merchant_returns_none(self):
        assert self._p(["15/03/2024", "", "99.99"]) is None

    def test_whitespace_only_merchant_returns_none(self):
        assert self._p(["15/03/2024", "   ", "99.99"]) is None

    def test_none_merchant_returns_none(self):
        assert self._p(["15/03/2024", None, "99.99"]) is None

    def test_row_shorter_than_expected_returns_none(self):
        # amount expected at col 2 but row only has 2 cells
        assert self._p(["15/03/2024", "Amazon"]) is None

    def test_non_standard_column_order(self):
        # amount=0, merchant=1, date=2
        cells = ["100.00", "Amazon", "15/03/2024"]
        tx = parse_transaction_row(cells, 2, 1, 0, _SRC, _SHEET)
        assert tx is not None
        assert tx.amount == Decimal("100.00")
        assert tx.date == date(2024, 3, 15)

    def test_normalized_merchant_differs_from_raw(self):
        tx = self._p(["15/03/2024", "AMAZON.COM, INC.", "99.99"])
        assert tx is not None
        assert tx.merchant == "AMAZON.COM, INC."
        assert tx.normalized_merchant == "amazoncom inc"

    def test_negative_amount_preserved(self):
        tx = self._p(["15/03/2024", "Refund Co", "-50.00"])
        assert tx is not None
        assert tx.amount == Decimal("-50.00")

    def test_parenthesised_negative_amount(self):
        tx = self._p(["15/03/2024", "Shop", "(30.00)"])
        assert tx is not None
        assert tx.amount == Decimal("-30.00")

    def test_amount_with_currency_symbol(self):
        tx = self._p(["15/03/2024", "Shop", "₪99.00"])
        assert tx is not None
        assert tx.amount == Decimal("99.00")

    def test_numeric_merchant_coerced_to_string(self):
        # Some bank files use numeric store codes as the merchant value.
        tx = self._p(["15/03/2024", 12345, "99.99"])
        assert tx is not None
        assert tx.merchant == "12345"


# ------------------------------------------------------------------ #
# _extract_from_sheet (pure state-machine)                            #
# ------------------------------------------------------------------ #

class TestExtractFromSheet:
    def _x(self, rows, config=None):
        return _extract_from_sheet(rows, _SRC, _SHEET, config or Config())

    def _x_stats(self, rows, config=None):
        return _extract_from_sheet_with_stats(rows, _SRC, _SHEET, config or Config())

    def test_simple_two_row_table(self):
        rows = [
            ["Purchase Date", "Merchant", "Amount"],
            ["15/03/2024", "Amazon", "99.99"],
            ["16/03/2024", "Walmart", "45.00"],
        ]
        result = self._x(rows)
        assert len(result) == 2
        assert result[0].merchant == "Amazon"
        assert result[1].merchant == "Walmart"

    def test_no_header_returns_empty(self):
        rows = [["15/03/2024", "Amazon", "99.99"]]
        assert self._x(rows) == []

    def test_header_only_sheet_returns_empty(self):
        rows = [["Purchase Date", "Merchant", "Amount"]]
        assert self._x(rows) == []

    def test_empty_sheet_returns_empty(self):
        assert self._x([]) == []

    def test_metadata_rows_before_header_ignored(self):
        rows = [
            ["Account: 123456"],
            ["Statement period: March 2024"],
            [""],
            ["Purchase Date", "Merchant", "Amount"],
            ["15/03/2024", "Amazon", "99.99"],
        ]
        result = self._x(rows)
        assert len(result) == 1
        assert result[0].merchant == "Amazon"

    def test_stop_token_ends_table(self):
        rows = [
            ["Purchase Date", "Merchant", "Amount"],
            ["15/03/2024", "Amazon", "99.99"],
            ["Total", None, "99.99"],
            ["16/03/2024", "Walmart", "45.00"],  # after stop — not extracted
        ]
        result = self._x(rows)
        assert len(result) == 1
        assert result[0].merchant == "Amazon"

    def test_all_empty_row_ends_table(self):
        rows = [
            ["Purchase Date", "Merchant", "Amount"],
            ["15/03/2024", "Amazon", "99.99"],
            [None, None, None],
            ["16/03/2024", "Walmart", "45.00"],  # no new header → not extracted
        ]
        result = self._x(rows)
        assert len(result) == 1

    def test_repeated_header_restarts_table(self):
        rows = [
            ["Purchase Date", "Merchant", "Amount"],
            ["15/03/2024", "Amazon", "99.99"],
            ["Purchase Date", "Merchant", "Amount"],  # second table
            ["16/03/2024", "Walmart", "45.00"],
        ]
        result = self._x(rows)
        assert len(result) == 2
        assert result[0].merchant == "Amazon"
        assert result[1].merchant == "Walmart"

    def test_two_tables_separated_by_stop_and_metadata(self):
        rows = [
            ["Purchase Date", "Merchant", "Amount"],
            ["15/03/2024", "Amazon", "99.99"],
            ["Total", None, "99.99"],
            ["Credit card transactions"],
            ["Purchase Date", "Merchant", "Amount"],
            ["16/03/2024", "Spotify", "9.99"],
        ]
        result = self._x(rows)
        assert len(result) == 2
        assert {tx.merchant for tx in result} == {"Amazon", "Spotify"}

    def test_invalid_date_row_skipped(self):
        rows = [
            ["Purchase Date", "Merchant", "Amount"],
            ["not-a-date", "Amazon", "99.99"],
            ["16/03/2024", "Walmart", "45.00"],
        ]
        result = self._x(rows)
        assert len(result) == 1
        assert result[0].merchant == "Walmart"

    def test_invalid_date_skip_count_recorded(self):
        rows = [
            ["Purchase Date", "Merchant", "Amount"],
            ["not-a-date", "Amazon", "99.99"],
            ["16/03/2024", "Walmart", "45.00"],
        ]
        result, stats = self._x_stats(rows)
        assert len(result) == 1
        assert isinstance(stats, ExtractionStats)
        assert stats.extracted_count == 1
        assert stats.skipped_invalid_date == 1

    def test_invalid_amount_row_skipped(self):
        rows = [
            ["Purchase Date", "Merchant", "Amount"],
            ["15/03/2024", "Amazon", "N/A"],
            ["16/03/2024", "Walmart", "45.00"],
        ]
        result = self._x(rows)
        assert len(result) == 1
        assert result[0].merchant == "Walmart"

    def test_source_metadata_preserved(self):
        rows = [
            ["Purchase Date", "Merchant", "Amount"],
            ["15/03/2024", "Amazon", "99.99"],
        ]
        result = _extract_from_sheet(rows, Path("bank_a.xlsx"), "March", Config())
        assert result[0].source_file == Path("bank_a.xlsx")
        assert result[0].sheet_name == "March"

    def test_hebrew_header_and_data(self):
        rows = [
            ["תאריך עסקה", "שם בית עסק", "סכום חיוב"],
            ["15/03/2024", "סופר פארם", "120.50"],
        ]
        result = self._x(rows)
        assert len(result) == 1
        assert result[0].merchant == "סופר פארם"
        assert result[0].amount == Decimal("120.50")

    def test_column_order_date_last(self):
        # Merchant | Amount | Date
        rows = [
            ["Merchant", "Amount", "Purchase Date"],
            ["Amazon", "99.99", "15/03/2024"],
        ]
        result = self._x(rows)
        assert len(result) == 1
        assert result[0].date == date(2024, 3, 15)
        assert result[0].amount == Decimal("99.99")

    def test_charge_amount_used_as_effective_amount_for_foreign_transaction(self):
        rows = [
            ["תאריך עסקה", "שם בית עסק", "סכום עסקה", "מטבע עסקה", "סכום חיוב", "מטבע חיוב"],
            ["17.03.26", "CLAUDE.AI SUBSCRIPTION", "20.00", "$", "73.40", "₪"],
        ]
        result = self._x(rows)
        assert len(result) == 1
        assert result[0].amount == Decimal("73.40")

    def test_charge_amount_is_converted_when_charge_currency_is_foreign(self, monkeypatch):
        monkeypatch.setattr(
            "ledgerflow.extraction.convert_amount_to_ils",
            lambda amount, currency_code, conversion_date, config: Decimal("73.40"),
        )
        rows = [
            ["תאריך עסקה", "שם בית עסק", "סכום עסקה", "מטבע עסקה", "סכום חיוב", "מטבע חיוב"],
            ["17.03.26", "CLAUDE.AI SUBSCRIPTION", "20.00", "$", "20.00", "$"],
        ]
        result = self._x(rows)
        assert len(result) == 1
        assert result[0].amount == Decimal("73.40")

    def test_foreign_charge_amount_is_converted_to_ils_using_charge_date(self, monkeypatch):
        captured: dict[str, object] = {}

        def fake_convert(amount, currency_code, conversion_date, config):
            captured["amount"] = amount
            captured["currency_code"] = currency_code
            captured["conversion_date"] = conversion_date
            captured["config"] = config
            return Decimal("73.40")

        monkeypatch.setattr("ledgerflow.extraction.convert_amount_to_ils", fake_convert)
        rows = [
            [
                "תאריך עסקה",
                "שם בית עסק",
                "סכום עסקה",
                "מטבע עסקה",
                "סכום חיוב",
                "מטבע חיוב",
                "חיוב בחשבון הבנק",
            ],
            ["17.03.26", "CLAUDE.AI SUBSCRIPTION", "20.00", "$", "20.00", "$", "19.03.26"],
        ]
        result = self._x(rows)
        assert len(result) == 1
        assert result[0].amount == Decimal("73.40")
        assert captured["amount"] == Decimal("20.00")
        assert captured["currency_code"] == "USD"
        assert captured["conversion_date"] == date(2026, 3, 19)

    def test_foreign_charge_amount_uses_purchase_date_when_charge_date_missing(self, monkeypatch):
        captured: dict[str, object] = {}

        def fake_convert(amount, currency_code, conversion_date, config):
            captured["conversion_date"] = conversion_date
            return Decimal("71.10")

        monkeypatch.setattr("ledgerflow.extraction.convert_amount_to_ils", fake_convert)
        rows = [
            ["תאריך עסקה", "שם בית עסק", "סכום עסקה", "מטבע עסקה", "סכום חיוב", "מטבע חיוב"],
            ["17.03.26", "CLAUDE.AI SUBSCRIPTION", "20.00", "$", "20.00", "$"],
        ]
        result = self._x(rows)
        assert len(result) == 1
        assert result[0].amount == Decimal("71.10")
        assert captured["conversion_date"] == date(2026, 3, 17)

    def test_foreign_converted_row_still_creates_exactly_one_transaction(self, monkeypatch):
        monkeypatch.setattr(
            "ledgerflow.extraction.convert_amount_to_ils",
            lambda amount, currency_code, conversion_date, config: Decimal("73.40"),
        )
        rows = [
            [
                "תאריך עסקה",
                "שם בית עסק",
                "סכום עסקה",
                "מטבע עסקה",
                "סכום חיוב",
                "מטבע חיוב",
                "חיוב בחשבון הבנק",
            ],
            ["17.03.26", "CLAUDE.AI SUBSCRIPTION", "20.00", "$", "20.00", "$", "19.03.26"],
        ]
        result = self._x(rows)
        assert len(result) == 1
        assert [tx.amount for tx in result] == [Decimal("73.40")]

    def test_regular_ils_transaction_without_charge_amount_remains_unchanged(self):
        rows = [
            ["תאריך עסקה", "שם בית עסק", "סכום עסקה", "מטבע עסקה"],
            ["17.03.26", "Supermarket", "120.50", "₪"],
        ]
        result = self._x(rows)
        assert len(result) == 1
        assert result[0].amount == Decimal("120.50")

    def test_installment_transaction_uses_monthly_charge_amount(self):
        rows = [
            [
                "תאריך עסקה",
                "שם בית עסק",
                "סכום עסקה",
                "מטבע עסקה",
                "סכום חיוב",
                "מטבע חיוב",
                "פירוט נוסף",
            ],
            [
                "20.01.26",
                "שטראוס מים בע\"מ",
                "5,532.01",
                "₪",
                "153.00",
                "₪",
                "תשלום 3 מתוך 36",
            ],
        ]
        result = self._x(rows)
        assert len(result) == 1
        assert result[0].amount == Decimal("153.00")

    def test_installment_transaction_creates_only_one_monthly_entry(self):
        rows = [
            [
                "תאריך עסקה",
                "שם בית עסק",
                "סכום עסקה",
                "מטבע עסקה",
                "סכום חיוב",
                "מטבע חיוב",
                "פירוט נוסף",
            ],
            [
                "20.01.26",
                "שטראוס מים בע\"מ",
                "5,532.01",
                "₪",
                "153.00",
                "₪",
                "תשלום 3 מתוך 36",
            ],
        ]
        result = self._x(rows)
        assert len(result) == 1
        assert [tx.amount for tx in result] == [Decimal("153.00")]

    def test_installment_text_does_not_cause_full_amount_usage(self):
        rows = [
            [
                "תאריך עסקה",
                "שם בית עסק",
                "סכום עסקה",
                "מטבע עסקה",
                "סכום חיוב",
                "מטבע חיוב",
                "פירוט נוסף",
            ],
            [
                "20.01.26",
                "חנות לדוגמה",
                "1,200.00",
                "₪",
                "99.90",
                "₪",
                "תשלום 1 מתוך 12",
            ],
        ]
        result = self._x(rows)
        assert len(result) == 1
        assert result[0].amount == Decimal("99.90")

    def test_installment_text_does_not_override_charge_amount_header_variant(self):
        rows = [
            [
                "תאריך עסקה",
                "שם בית עסק",
                "סכום עסקה",
                "מטבע עסקה",
                "סכום חיוב בש\"ח",
                "מטבע חיוב",
                "פירוט נוסף",
            ],
            [
                "20.01.26",
                "שטראוס מים בע\"מ",
                "5,532.01",
                "₪",
                "153.00",
                "₪",
                "תשלום 3 מתוך 36",
            ],
        ]
        result = self._x(rows)
        assert len(result) == 1
        assert result[0].amount == Decimal("153.00")

    def test_installment_details_are_preserved_for_month_filter_override(self):
        rows = [
            [
                "תאריך עסקה",
                "שם בית עסק",
                "סכום עסקה",
                "מטבע עסקה",
                "סכום חיוב",
                "מטבע חיוב",
                "פירוט נוסף",
            ],
            [
                "20.01.26",
                "שטראוס מים בע\"מ",
                "5,532.01",
                "₪",
                "153.00",
                "₪",
                "תשלום 3 מתוך 36",
            ],
        ]
        result = self._x(rows)
        assert len(result) == 1
        assert result[0].additional_details == "תשלום 3 מתוך 36"


# ------------------------------------------------------------------ #
# extract_from_file (file-based integration)                           #
# ------------------------------------------------------------------ #

class TestExtractFromFile:
    def test_xls_extension_raises_unsupported(self, tmp_path):
        path = tmp_path / "bank.xls"
        path.write_bytes(b"fake content")
        with pytest.raises(UnsupportedFileTypeError, match=r"\.xls"):
            extract_from_file(path, Config())

    def test_unsupported_extension_raises(self, tmp_path):
        path = tmp_path / "bank.csv"
        path.write_bytes(b"a,b,c")
        with pytest.raises(UnsupportedFileTypeError):
            extract_from_file(path, Config())

    def test_xlsb_extension_raises(self, tmp_path):
        path = tmp_path / "bank.xlsb"
        path.write_bytes(b"fake")
        with pytest.raises(UnsupportedFileTypeError):
            extract_from_file(path, Config())

    def test_missing_file_raises(self, tmp_path):
        with pytest.raises(Exception):
            extract_from_file(tmp_path / "nonexistent.xlsx", Config())

    def test_single_sheet_extraction(self, tmp_path):
        path = write_workbook(tmp_path, {
            "Transactions": [
                ["Purchase Date", "Merchant", "Amount"],
                [datetime(2024, 3, 15), "Amazon", 99.99],
                [datetime(2024, 3, 16), "Walmart", 45.00],
            ]
        })
        result = extract_from_file(path, Config())
        assert len(result) == 2
        assert result[0].merchant == "Amazon"
        assert result[1].merchant == "Walmart"

    def test_extract_from_file_with_stats(self, tmp_path):
        path = write_workbook(tmp_path, {
            "Transactions": [
                ["Charge Date", "Merchant", "Amount"],
                ["28.03.26", "Amazon", "99.99"],
                ["bad-date", "Walmart", "45.00"],
            ]
        })
        result, stats = extract_from_file_with_stats(path, Config())
        assert len(result) == 1
        assert result[0].date == date(2026, 3, 28)
        assert stats.header_rows_detected == 1
        assert stats.extracted_count == 1
        assert stats.skipped_invalid_date == 1

    def test_multiple_sheets_all_extracted(self, tmp_path):
        path = write_workbook(tmp_path, {
            "Sheet1": [
                ["Purchase Date", "Merchant", "Amount"],
                [datetime(2024, 3, 15), "Amazon", 99.99],
            ],
            "Sheet2": [
                ["Purchase Date", "Merchant", "Amount"],
                [datetime(2024, 3, 16), "Walmart", 45.00],
            ],
        })
        result = extract_from_file(path, Config())
        assert len(result) == 2
        assert {tx.merchant for tx in result} == {"Amazon", "Walmart"}

    def test_sheet_without_table_contributes_nothing(self, tmp_path):
        path = write_workbook(tmp_path, {
            "Summary": [["Account: John"], ["Period: March 2024"]],
            "Transactions": [
                ["Purchase Date", "Merchant", "Amount"],
                [datetime(2024, 3, 15), "Amazon", 99.99],
            ],
        })
        result = extract_from_file(path, Config())
        assert len(result) == 1

    def test_source_file_path_preserved(self, tmp_path):
        path = write_workbook(tmp_path, {
            "Transactions": [
                ["Purchase Date", "Merchant", "Amount"],
                [datetime(2024, 3, 15), "Amazon", 99.99],
            ]
        })
        result = extract_from_file(path, Config())
        assert result[0].source_file == path

    def test_sheet_name_preserved(self, tmp_path):
        path = write_workbook(tmp_path, {
            "March 2024": [
                ["Purchase Date", "Merchant", "Amount"],
                [datetime(2024, 3, 15), "Amazon", 99.99],
            ]
        })
        result = extract_from_file(path, Config())
        assert result[0].sheet_name == "March 2024"

    def test_dates_from_real_datetime_cells(self, tmp_path):
        path = write_workbook(tmp_path, {
            "Data": [
                ["Purchase Date", "Merchant", "Amount"],
                [datetime(2024, 6, 1, 0, 0), "Spotify", 9.99],
            ]
        })
        result = extract_from_file(path, Config())
        assert result[0].date == date(2024, 6, 1)

    def test_xlsm_extension_accepted(self, tmp_path):
        # openpyxl can open a structurally-valid xlsx saved with a .xlsm extension.
        xlsx_path = write_workbook(tmp_path, {
            "Transactions": [
                ["Purchase Date", "Merchant", "Amount"],
                [datetime(2024, 3, 15), "Amazon", 99.99],
            ]
        })
        xlsm_path = tmp_path / "bank.xlsm"
        xlsx_path.rename(xlsm_path)
        result = extract_from_file(xlsm_path, Config())
        assert len(result) == 1

    def test_stop_row_in_real_workbook(self, tmp_path):
        path = write_workbook(tmp_path, {
            "Data": [
                ["Purchase Date", "Merchant", "Amount"],
                [datetime(2024, 3, 15), "Amazon", 99.99],
                ["Total", None, 99.99],
                [datetime(2024, 3, 16), "Walmart", 45.00],  # after stop
            ]
        })
        result = extract_from_file(path, Config())
        assert len(result) == 1
        assert result[0].merchant == "Amazon"


# ------------------------------------------------------------------ #
# extract_from_files                                                   #
# ------------------------------------------------------------------ #

class TestExtractFromFiles:
    def test_combines_results_from_two_files(self, tmp_path):
        path_a = write_workbook(tmp_path, {
            "Transactions": [
                ["Purchase Date", "Merchant", "Amount"],
                [datetime(2024, 3, 15), "Amazon", 99.99],
            ]
        }, name="bank_a.xlsx")
        path_b = write_workbook(tmp_path, {
            "Transactions": [
                ["Purchase Date", "Merchant", "Amount"],
                [datetime(2024, 3, 16), "Walmart", 45.00],
            ]
        }, name="bank_b.xlsx")

        result = extract_from_files([path_a, path_b], Config())
        assert len(result) == 2
        assert {tx.merchant for tx in result} == {"Amazon", "Walmart"}

    def test_empty_list_returns_empty(self):
        assert extract_from_files([], Config()) == []

    def test_extract_from_files_with_stats(self, tmp_path):
        path_a = write_workbook(tmp_path, {
            "Transactions": [
                ["Charge Date", "Merchant", "Amount"],
                ["28.03.26", "Amazon", "99.99"],
            ]
        }, name="bank_a.xlsx")
        path_b = write_workbook(tmp_path, {
            "Transactions": [
                ["Purchase Date", "Merchant", "Amount"],
                ["bad-date", "Walmart", "45.00"],
            ]
        }, name="bank_b.xlsx")

        result, stats = extract_from_files_with_stats([path_a, path_b], Config())
        assert len(result) == 1
        assert stats.extracted_count == 1
        assert stats.skipped_invalid_date == 1

    def test_source_file_distinguishes_origin(self, tmp_path):
        path_a = write_workbook(tmp_path, {
            "Transactions": [
                ["Purchase Date", "Merchant", "Amount"],
                [datetime(2024, 3, 15), "Amazon", 99.99],
            ]
        }, name="bank_a.xlsx")
        path_b = write_workbook(tmp_path, {
            "Transactions": [
                ["Purchase Date", "Merchant", "Amount"],
                [datetime(2024, 3, 16), "Walmart", 45.00],
            ]
        }, name="bank_b.xlsx")

        result = extract_from_files([path_a, path_b], Config())
        sources = {tx.source_file.name for tx in result}
        assert sources == {"bank_a.xlsx", "bank_b.xlsx"}
