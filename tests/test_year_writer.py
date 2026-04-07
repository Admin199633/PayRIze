"""Tests for ledgerflow.year_writer."""
from __future__ import annotations

import time
from datetime import date
from decimal import Decimal
from pathlib import Path

import openpyxl
import pytest

from ledgerflow.config import Config
from ledgerflow.models import Category, Transaction
from ledgerflow.year_writer import (
    CATEGORY_ALIASES,
    FIRST_DATA_ROW_INDEX,
    CategoryBlock,
    CategoryBlockNotFoundError,
    MonthSheetNotFoundError,
    WriteResult,
    create_backup,
    detect_category_blocks,
    find_first_empty_row,
    resolve_month_sheet,
    write_yearly_workbook,
)


# ------------------------------------------------------------------ #
# Helpers                                                              #
# ------------------------------------------------------------------ #


def make_config(include_mortgage: bool = False) -> Config:
    cfg = Config()
    cfg.include_mortgage = include_mortgage
    if include_mortgage and Category.MORTGAGE not in cfg.categories:
        cfg.categories.append(Category.MORTGAGE)
    return cfg


def tx(
    merchant: str = "Shop",
    amount: str = "10.00",
    d: str = "2024-03-15",
    category: Category = Category.OTHER,
) -> Transaction:
    return Transaction(
        date=date.fromisoformat(d),
        merchant=merchant,
        normalized_merchant=merchant.lower(),
        amount=Decimal(amount),
        source_file=Path("bank.xlsx"),
        sheet_name="Sheet1",
        category=category,
    )


def make_workbook(tmp_path: Path, sheet_name: str = "March") -> Path:
    """
    Minimal workbook with one sheet containing a single Other block.
    Category label at A1; the default layout (col 0) gives merchant=A, price=B, date=C.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws["A1"] = "Other"
    ws["B1"] = "Price"
    ws["C1"] = "Date"
    path = tmp_path / "yearly.xlsx"
    wb.save(str(path))
    return path


def make_workbook_multi_block(tmp_path: Path, sheet_name: str = "March") -> Path:
    """
    Two category blocks using the shifted layout (category label is in the price column):
      Leisure  at col C (index 2) → merchant=B(1), price=C(2), date=D(3)
      Business at col G (index 6) → merchant=F(5), price=G(6), date=H(7)
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws["C1"] = "Leisure"
    ws["D1"] = "Price"
    ws["E1"] = "Date"
    ws["G1"] = "Business"
    ws["H1"] = "Price"
    ws["I1"] = "Date"
    path = tmp_path / "yearly.xlsx"
    wb.save(str(path))
    return path


def read_sheet_values(path: Path, sheet_name: str) -> list[tuple]:
    wb = openpyxl.load_workbook(str(path))
    ws = wb[sheet_name]
    return [tuple(row) for row in ws.iter_rows(values_only=True)]


# ------------------------------------------------------------------ #
# create_backup                                                        #
# ------------------------------------------------------------------ #


class TestCreateBackup:
    def test_backup_file_exists(self, tmp_path):
        src = make_workbook(tmp_path)
        backup = create_backup(src)
        assert backup.exists()

    def test_backup_is_different_path_from_source(self, tmp_path):
        src = make_workbook(tmp_path)
        backup = create_backup(src)
        assert backup.resolve() != src.resolve()

    def test_backup_contains_timestamp(self, tmp_path):
        src = make_workbook(tmp_path)
        backup = create_backup(src)
        # Stem format: original_YYYYMMDD_HHMMSS
        parts = backup.stem.split("_")
        assert len(parts) >= 3

    def test_backup_same_suffix_as_source(self, tmp_path):
        src = make_workbook(tmp_path)
        backup = create_backup(src)
        assert backup.suffix == src.suffix

    def test_backup_is_readable_workbook(self, tmp_path):
        src = make_workbook(tmp_path)
        backup = create_backup(src)
        wb = openpyxl.load_workbook(str(backup))
        assert wb is not None

    def test_backup_in_custom_dir(self, tmp_path):
        src = make_workbook(tmp_path)
        backup_dir = tmp_path / "backups"
        backup = create_backup(src, backup_dir)
        assert backup.parent.resolve() == backup_dir.resolve()
        assert backup.exists()

    def test_backup_defaults_to_source_dir(self, tmp_path):
        src = make_workbook(tmp_path)
        backup = create_backup(src)
        assert backup.parent.resolve() == src.parent.resolve()

    def test_source_file_unchanged_after_backup(self, tmp_path):
        src = make_workbook(tmp_path)
        original_size = src.stat().st_size
        create_backup(src)
        assert src.stat().st_size == original_size

    def test_two_backups_have_distinct_paths(self, tmp_path):
        src = make_workbook(tmp_path)
        b1 = create_backup(src)
        time.sleep(1.1)  # ensure different timestamp
        b2 = create_backup(src)
        assert b1 != b2


# ------------------------------------------------------------------ #
# resolve_month_sheet                                                  #
# ------------------------------------------------------------------ #


class TestResolveMonthSheet:
    def test_exact_zero_padded_number(self):
        assert resolve_month_sheet(["01", "02", "03"], 1) == "01"

    def test_exact_bare_number(self):
        assert resolve_month_sheet(["1", "2", "12"], 12) == "12"

    def test_exact_english_name(self):
        assert resolve_month_sheet(["January", "February"], 1) == "January"

    def test_exact_english_abbreviation(self):
        assert resolve_month_sheet(["Jan", "Feb", "Mar"], 3) == "Mar"

    def test_exact_hebrew_name(self):
        assert resolve_month_sheet(["ינואר", "פברואר", "מרץ"], 1) == "ינואר"

    def test_case_insensitive_english(self):
        assert resolve_month_sheet(["january", "FEBRUARY"], 2) == "FEBRUARY"

    def test_hebrew_month_3(self):
        assert resolve_month_sheet(["מרץ"], 3) == "מרץ"

    def test_all_twelve_hebrew_names(self):
        hebrew = {
            1: "ינואר", 2: "פברואר", 3: "מרץ", 4: "אפריל",
            5: "מאי", 6: "יוני", 7: "יולי", 8: "אוגוסט",
            9: "ספטמבר", 10: "אוקטובר", 11: "נובמבר", 12: "דצמבר",
        }
        names = list(hebrew.values())
        for m, name in hebrew.items():
            assert resolve_month_sheet(names, m) == name

    def test_sheet_with_leading_trailing_spaces(self):
        assert resolve_month_sheet(["  March  "], 3) == "  March  "

    def test_contains_english_name(self):
        assert resolve_month_sheet(["Budget_March_2024"], 3) == "Budget_March_2024"

    def test_contains_number(self):
        assert resolve_month_sheet(["Month_03"], 3) == "Month_03"

    def test_exact_match_preferred_over_contains(self):
        assert resolve_month_sheet(["March2024", "March"], 3) == "March"

    def test_no_match_raises(self):
        # "Summary" contains "mar" as a substring so cannot be used here.
        with pytest.raises(MonthSheetNotFoundError):
            resolve_month_sheet(["Budget", "Totals"], 3)

    def test_invalid_month_raises_value_error(self):
        with pytest.raises(ValueError):
            resolve_month_sheet(["Sheet1"], 0)

    def test_invalid_month_13_raises(self):
        with pytest.raises(ValueError):
            resolve_month_sheet(["Sheet1"], 13)

    def test_empty_sheet_list_raises(self):
        with pytest.raises(MonthSheetNotFoundError):
            resolve_month_sheet([], 1)

    def test_no_silent_fallback_to_active_sheet(self):
        with pytest.raises(MonthSheetNotFoundError):
            resolve_month_sheet(["Summary", "Data", "Config"], 5)


# ------------------------------------------------------------------ #
# detect_category_blocks                                               #
# ------------------------------------------------------------------ #


class TestDetectCategoryBlocks:
    def test_detects_english_other_at_col_0(self):
        rows = [("Other", None, None)]
        blocks = detect_category_blocks(rows, make_config())
        assert Category.OTHER in blocks

    def test_col_0_uses_simple_layout(self):
        """Category at col 0: merchant=0, price=1, date=2 (no left-shift)."""
        rows = [("Other", None, None)]
        blocks = detect_category_blocks(rows, make_config())
        b = blocks[Category.OTHER]
        assert b.merchant_col == 0
        assert b.price_col == 1
        assert b.date_col == 2

    def test_col_nonzero_uses_shifted_layout(self):
        """Category at col 2: merchant=1, price=2, date=3 (shifted left by 1)."""
        rows = [(None, None, "Business", None, None)]
        blocks = detect_category_blocks(rows, make_config())
        b = blocks[Category.BUSINESS]
        assert b.merchant_col == 1
        assert b.price_col == 2
        assert b.date_col == 3

    def test_first_data_row_is_always_first_data_row_index(self):
        rows = [("Other", None, None)]
        blocks = detect_category_blocks(rows, make_config())
        assert blocks[Category.OTHER].first_data_row == FIRST_DATA_ROW_INDEX

    def test_triplet_subheader_shifts_columns(self):
        """
        When a merchant/price/date header triplet appears in the next row to
        the RIGHT of the category anchor, columns are resolved from it.
        """
        rows = [
            (None, "Other", None, None, None),
            (None, None, "פירוט עסקה", "מחיר", "תאריך"),
        ]
        blocks = detect_category_blocks(rows, make_config())
        b = blocks[Category.OTHER]
        # Triplet at cols 2,3,4 → shift left by 1 → merchant=1, price=2, date=3
        assert b.merchant_col == 1
        assert b.price_col == 2
        assert b.date_col == 3

    def test_detects_hebrew_category(self):
        rows = [("ביטוח", None, None)]
        blocks = detect_category_blocks(rows, make_config())
        assert Category.INSURANCE in blocks

    def test_case_insensitive_header(self):
        rows = [("LEISURE", None, None)]
        blocks = detect_category_blocks(rows, make_config())
        assert Category.LEISURE in blocks

    def test_header_with_whitespace(self):
        rows = [(None, "  Home  ", None, None)]
        blocks = detect_category_blocks(rows, make_config())
        assert Category.HOME in blocks

    def test_multiple_blocks_same_row(self):
        row = [None] * 10
        row[0] = "Leisure"
        row[5] = "Business"
        blocks = detect_category_blocks([tuple(row)], make_config())
        assert Category.LEISURE in blocks
        assert Category.BUSINESS in blocks

    def test_multiple_blocks_different_rows(self):
        row1 = ("Other", None, None)
        row2 = ("Business", None, None)
        blocks = detect_category_blocks([row1, row2], make_config())
        assert Category.OTHER in blocks
        assert Category.BUSINESS in blocks

    def test_first_occurrence_wins(self):
        # "Other" appears at row 0 col 0 and at row 1 col 2
        rows = [
            ("Other", None, None, None),
            (None, None, "Other", None),
        ]
        blocks = detect_category_blocks(rows, make_config())
        assert blocks[Category.OTHER].header_row == 0

    def test_unknown_header_ignored(self):
        rows = [("Unknown Category", None, None)]
        assert detect_category_blocks(rows, make_config()) == {}

    def test_mortgage_excluded_when_not_allowed(self):
        rows = [("Mortgage", None, None)]
        blocks = detect_category_blocks(rows, make_config(include_mortgage=False))
        assert Category.MORTGAGE not in blocks

    def test_mortgage_included_when_allowed(self):
        rows = [("Mortgage", None, None)]
        blocks = detect_category_blocks(rows, make_config(include_mortgage=True))
        assert Category.MORTGAGE in blocks

    def test_empty_rows(self):
        assert detect_category_blocks([], make_config()) == {}

    def test_all_none_row(self):
        assert detect_category_blocks([(None, None, None)], make_config()) == {}

    def test_header_row_index_recorded(self):
        rows = [(None,), ("Other", None, None)]
        blocks = detect_category_blocks(rows, make_config())
        assert blocks[Category.OTHER].header_row == 1


# ------------------------------------------------------------------ #
# find_first_empty_row                                                 #
# ------------------------------------------------------------------ #


class TestFindFirstEmptyRow:
    def _block(self, header_row: int = 0, col: int = 0) -> CategoryBlock:
        return CategoryBlock(
            category=Category.OTHER,
            header_row=header_row,
            first_data_row=FIRST_DATA_ROW_INDEX,
            merchant_col=col,
            price_col=col + 1,
            date_col=col + 2,
        )

    def test_always_starts_at_first_data_row_index(self):
        """Row 7 (index 6) is always the starting point regardless of header position."""
        # Rows 0-5 are all empty, row 6 (Excel row 7) is also empty → returns 6.
        rows = [(None, None, None)] * 7
        assert find_first_empty_row(rows, self._block(header_row=0)) == FIRST_DATA_ROW_INDEX

    def test_placeholder_price_at_row_7_is_writable(self):
        """
        A placeholder value (e.g. "1") in the price cell of row 7 must NOT
        block writing.  Only the merchant cell determines occupancy.
        """
        rows = [(None, None, None)] * 7
        # Row 6 (Excel row 7): merchant=None, price="1" (template placeholder), date=None
        rows[6] = (None, "1", None)
        assert find_first_empty_row(rows, self._block()) == FIRST_DATA_ROW_INDEX

    def test_numeric_placeholder_in_price_is_writable(self):
        """A numeric 1 (not a string) in the price cell is also writable."""
        rows = [(None, None, None)] * 7
        rows[6] = (None, 1, None)
        assert find_first_empty_row(rows, self._block()) == FIRST_DATA_ROW_INDEX

    def test_real_merchant_at_row_7_is_occupied(self):
        """A real merchant name at row 7 means the row is taken; returns row 8."""
        rows = [(None, None, None)] * 7
        rows[6] = ("Shop A", 55.0, date(2024, 3, 1))
        assert find_first_empty_row(rows, self._block()) == 7

    def test_skips_occupied_rows_and_finds_next_empty(self):
        rows = [(None, None, None)] * 9
        rows[6] = ("Shop A", 10.0, date(2024, 3, 1))
        rows[7] = ("Shop B", 20.0, date(2024, 3, 2))
        assert find_first_empty_row(rows, self._block()) == 8

    def test_all_occupied_returns_past_last_row(self):
        rows = [(None, None, None)] * 7
        rows[6] = ("Shop A", 10.0, date(2024, 3, 1))
        assert find_first_empty_row(rows, self._block()) == 7

    def test_whitespace_merchant_counts_as_empty(self):
        rows = [(None, None, None)] * 7
        rows[6] = ("   ", "1", "\t")
        assert find_first_empty_row(rows, self._block()) == FIRST_DATA_ROW_INDEX

    def test_col_beyond_row_length_counts_as_empty(self):
        rows = [(None, None, None)] * 7
        rows[6] = (None,)  # very short row; merchant col 0 is None
        assert find_first_empty_row(rows, self._block()) == FIRST_DATA_ROW_INDEX

    def test_offset_block(self):
        """Block at col 3 uses that column for merchant emptiness check."""
        rows = [(None, None, None, None, None, None)] * 7
        rows[6] = (None, None, None, "Business", 50.0, date(2024, 3, 1))
        block = self._block(header_row=0, col=3)
        # Row 6 has merchant at col 3 → occupied → returns 7
        assert find_first_empty_row(rows, block) == 7

    def test_returns_first_data_row_idx_when_rows_list_shorter(self):
        """If the sheet has fewer rows than first_data_row, returns first_data_row."""
        rows = [(None, None, None)] * 3  # only 3 rows
        assert find_first_empty_row(rows, self._block()) == FIRST_DATA_ROW_INDEX

    def test_price_only_in_row_7_treated_as_empty(self):
        """Row with only price set (no merchant) is still writable."""
        rows = [(None, None, None)] * 7
        rows[6] = (None, 50.0, None)
        assert find_first_empty_row(rows, self._block()) == FIRST_DATA_ROW_INDEX


# ------------------------------------------------------------------ #
# write_yearly_workbook — integration tests                            #
# ------------------------------------------------------------------ #


class TestWriteYearlyWorkbook:
    # ── safety contract ───────────────────────────────────────────────

    def test_raises_when_output_equals_source(self, tmp_path):
        src = make_workbook(tmp_path)
        with pytest.raises(ValueError, match="output_path must differ"):
            write_yearly_workbook([], src, src, 3, make_config())

    def test_backup_created_before_write(self, tmp_path):
        src = make_workbook(tmp_path)
        out = tmp_path / "out.xlsx"
        result = write_yearly_workbook([], src, out, 3, make_config())
        assert result.backup_path.exists()

    def test_source_workbook_unchanged(self, tmp_path):
        src = make_workbook(tmp_path)
        original_bytes = src.read_bytes()
        out = tmp_path / "out.xlsx"
        write_yearly_workbook([tx()], src, out, 3, make_config())
        assert src.read_bytes() == original_bytes

    def test_output_path_created(self, tmp_path):
        src = make_workbook(tmp_path)
        out = tmp_path / "result" / "out.xlsx"
        write_yearly_workbook([], src, out, 3, make_config())
        assert out.exists()

    def test_returns_write_result(self, tmp_path):
        src = make_workbook(tmp_path)
        out = tmp_path / "out.xlsx"
        result = write_yearly_workbook([], src, out, 3, make_config())
        assert isinstance(result, WriteResult)
        assert result.output_path == out

    # ── month sheet resolution ────────────────────────────────────────

    def test_raises_when_month_sheet_not_found(self, tmp_path):
        # Note: "Summary" contains the substring "mar" so it would match month 3
        # via the contains fallback. Use a name that is clearly unrelated.
        wb = openpyxl.Workbook()
        wb.active.title = "Budget"
        src = tmp_path / "yearly.xlsx"
        wb.save(str(src))
        out = tmp_path / "out.xlsx"
        with pytest.raises(MonthSheetNotFoundError):
            write_yearly_workbook([], src, out, 3, make_config())

    def test_uses_hebrew_month_sheet(self, tmp_path):
        src = make_workbook(tmp_path, sheet_name="מרץ")
        out = tmp_path / "out.xlsx"
        result = write_yearly_workbook([tx()], src, out, 3, make_config())
        assert result.transactions_written == 1

    # ── row 7 is always the first write row ───────────────────────────

    def test_first_transaction_written_to_row_7(self, tmp_path):
        """Transactions always land at Excel row 7 (0-indexed row 6)."""
        src = make_workbook(tmp_path)
        out = tmp_path / "out.xlsx"
        write_yearly_workbook([tx(merchant="Supermarket", amount="55.00")], src, out, 3, make_config())
        rows = read_sheet_values(out, "March")
        # Category label at A1 uses simple layout: merchant=A, price=B, date=C
        assert rows[FIRST_DATA_ROW_INDEX][0] == "Supermarket"
        assert rows[FIRST_DATA_ROW_INDEX][1] == 55.0

    def test_placeholder_in_price_at_row_7_does_not_push_to_row_8(self, tmp_path):
        """
        A template placeholder value such as '1' in the price cell of row 7
        must not cause the writer to skip to row 8.
        """
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "March"
        ws["A1"] = "Other"
        ws["B1"] = "Price"
        ws["C1"] = "Date"
        # Template placeholder in the price column of row 7
        ws["B7"] = 1
        src = tmp_path / "yearly.xlsx"
        wb.save(str(path := src))
        out = tmp_path / "out.xlsx"
        write_yearly_workbook([tx(merchant="Cinema", amount="35.00")], path, out, 3, make_config())
        rows = read_sheet_values(out, "March")
        # Must land at row 7 (index 6), not row 8 (index 7)
        assert rows[FIRST_DATA_ROW_INDEX][0] == "Cinema"
        assert rows[FIRST_DATA_ROW_INDEX][1] == 35.0
        assert len(rows) < 8 or rows[7][0] is None  # nothing written to row 8

    def test_category_label_in_row_1_untouched(self, tmp_path):
        """The category header cell must survive a full write pass."""
        src = make_workbook(tmp_path)
        out = tmp_path / "out.xlsx"
        write_yearly_workbook([tx()], src, out, 3, make_config())
        wb = openpyxl.load_workbook(str(out))
        ws = wb["March"]
        assert ws["A1"].value == "Other"

    def test_date_written_to_date_column(self, tmp_path):
        src = make_workbook(tmp_path)
        out = tmp_path / "out.xlsx"
        write_yearly_workbook([tx(d="2024-03-20")], src, out, 3, make_config())
        rows = read_sheet_values(out, "March")
        cell = rows[FIRST_DATA_ROW_INDEX][2]
        # openpyxl may return datetime.datetime for date cells
        assert getattr(cell, "year", None) == 2024
        assert getattr(cell, "month", None) == 3
        assert getattr(cell, "day", None) == 20

    def test_multiple_transactions_written_in_order(self, tmp_path):
        src = make_workbook(tmp_path)
        out = tmp_path / "out.xlsx"
        txs = [tx(merchant="A", amount="1.00"), tx(merchant="B", amount="2.00")]
        write_yearly_workbook(txs, src, out, 3, make_config())
        rows = read_sheet_values(out, "March")
        assert rows[FIRST_DATA_ROW_INDEX][0] == "A"
        assert rows[FIRST_DATA_ROW_INDEX + 1][0] == "B"

    def test_transactions_written_count_returned(self, tmp_path):
        src = make_workbook(tmp_path)
        out = tmp_path / "out.xlsx"
        result = write_yearly_workbook([tx(), tx(), tx()], src, out, 3, make_config())
        assert result.transactions_written == 3

    def test_zero_transactions_returns_zero_count(self, tmp_path):
        src = make_workbook(tmp_path)
        out = tmp_path / "out.xlsx"
        result = write_yearly_workbook([], src, out, 3, make_config())
        assert result.transactions_written == 0

    # ── category block routing ────────────────────────────────────────

    def test_transactions_routed_to_correct_block(self, tmp_path):
        """
        make_workbook_multi_block layout (shifted: category label is in price column):
          Leisure  at col C (2) → merchant=B(1), price=C(2), date=D(3)
          Business at col G (6) → merchant=F(5), price=G(6), date=H(7)
        """
        src = make_workbook_multi_block(tmp_path)
        out = tmp_path / "out.xlsx"
        txs = [
            tx(merchant="Cinema", category=Category.LEISURE),
            tx(merchant="Office", category=Category.BUSINESS),
        ]
        write_yearly_workbook(txs, src, out, 3, make_config())
        rows = read_sheet_values(out, "March")
        # Leisure merchant column = B (index 1)
        assert rows[FIRST_DATA_ROW_INDEX][1] == "Cinema"
        # Business merchant column = F (index 5)
        assert rows[FIRST_DATA_ROW_INDEX][5] == "Office"

    def test_fallback_to_other_block(self, tmp_path):
        """Transaction with Leisure category but only Other block → Other block."""
        src = make_workbook(tmp_path)
        out = tmp_path / "out.xlsx"
        write_yearly_workbook([tx(merchant="Cinema", category=Category.LEISURE)], src, out, 3, make_config())
        rows = read_sheet_values(out, "March")
        assert rows[FIRST_DATA_ROW_INDEX][0] == "Cinema"

    def test_raises_when_no_block_and_no_other(self, tmp_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "March"
        ws["B1"] = "Business"
        ws["C1"] = "Price"
        ws["D1"] = "Date"
        src = tmp_path / "yearly.xlsx"
        wb.save(str(src))
        out = tmp_path / "out.xlsx"
        with pytest.raises(CategoryBlockNotFoundError):
            write_yearly_workbook(
                [tx(merchant="Cinema", category=Category.LEISURE)],
                src, out, 3, make_config(),
            )

    def test_appends_below_existing_data(self, tmp_path):
        """Existing data in block must not be overwritten."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "March"
        ws["A1"] = "Other"
        ws["B1"] = "Price"
        ws["C1"] = "Date"
        ws["A7"] = "Existing"
        ws["B7"] = 99.0
        ws["C7"] = date(2024, 3, 1)
        src = tmp_path / "yearly.xlsx"
        wb.save(str(src))
        out = tmp_path / "out.xlsx"
        write_yearly_workbook([tx(merchant="New", amount="5.00")], src, out, 3, make_config())
        rows = read_sheet_values(out, "March")
        assert rows[FIRST_DATA_ROW_INDEX][0] == "Existing"
        assert rows[FIRST_DATA_ROW_INDEX + 1][0] == "New"

    # ── mortgage ─────────────────────────────────────────────────────

    def test_mortgage_block_ignored_when_not_allowed(self, tmp_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "March"
        ws["A1"] = "Other"
        ws["B1"] = "Price"
        ws["C1"] = "Date"
        ws["F1"] = "Mortgage"
        ws["G1"] = "Price"
        ws["H1"] = "Date"
        src = tmp_path / "yearly.xlsx"
        wb.save(str(src))
        out = tmp_path / "out.xlsx"
        write_yearly_workbook(
            [tx(merchant="Bank", category=Category.MORTGAGE)],
            src, out, 3, make_config(include_mortgage=False),
        )
        rows = read_sheet_values(out, "March")
        # Fell back to Other block: merchant at A7 (index 0)
        assert rows[FIRST_DATA_ROW_INDEX][0] == "Bank"

    def test_mortgage_block_used_when_allowed(self, tmp_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "March"
        ws["A1"] = "Other"
        ws["B1"] = "Price"
        ws["C1"] = "Date"
        ws["F1"] = "Mortgage"
        ws["G1"] = "Price"
        ws["H1"] = "Date"
        src = tmp_path / "yearly.xlsx"
        wb.save(str(src))
        out = tmp_path / "out.xlsx"
        write_yearly_workbook(
            [tx(merchant="Bank", category=Category.MORTGAGE)],
            src, out, 3, make_config(include_mortgage=True),
        )
        rows = read_sheet_values(out, "March")
        # Mortgage block: category at F (5) → merchant at E (4)
        assert rows[FIRST_DATA_ROW_INDEX][4] == "Bank"

    # ── backup dir ────────────────────────────────────────────────────

    def test_custom_backup_dir_used(self, tmp_path):
        src = make_workbook(tmp_path)
        out = tmp_path / "out.xlsx"
        backup_dir = tmp_path / "bk"
        result = write_yearly_workbook([], src, out, 3, make_config(), backup_dir=backup_dir)
        assert result.backup_path.parent.resolve() == backup_dir.resolve()


# ------------------------------------------------------------------ #
# CATEGORY_ALIASES completeness                                        #
# ------------------------------------------------------------------ #


class TestCategoryAliases:
    def test_all_non_mortgage_categories_have_english_alias(self):
        cats = [
            Category.OTHER, Category.BUSINESS, Category.PHARM,
            Category.INSURANCE, Category.TRANSPORTATION, Category.LEISURE,
            Category.HOME,
        ]
        for cat in cats:
            assert cat.value.lower() in CATEGORY_ALIASES, (
                f"Expected {cat.value.lower()!r} in CATEGORY_ALIASES for {cat}"
            )

    def test_mortgage_has_english_alias(self):
        assert "mortgage" in CATEGORY_ALIASES
        assert CATEGORY_ALIASES["mortgage"] == Category.MORTGAGE

    def test_all_aliases_map_to_valid_category(self):
        for key, value in CATEGORY_ALIASES.items():
            assert isinstance(value, Category), f"{key!r} maps to non-Category {value!r}"

    def test_no_empty_string_key(self):
        assert "" not in CATEGORY_ALIASES
