"""
Yearly workbook writer.

Writes final categorized transactions into a pre-existing yearly Excel workbook
that has one sheet per calendar month and one category block per category.

Safety contract
---------------
1. A timestamped backup is created from the source workbook BEFORE any
   modification.  If backup creation fails the entire write is aborted.
2. The source workbook is NEVER modified — all writes go to ``output_path``.
3. ``output_path`` must differ from ``source_workbook`` (validated at entry).

Expected workbook structure
---------------------------
Each month sheet contains one or more category blocks.  A "category block"
is a contiguous set of three adjacent columns whose first cell (the block
header) contains a recognisable category name:

    | <Category>  | <price header> | <date header> |
    | merchant 1  |  amount        |  date         |
    | merchant 2  |  amount        |  date         |
    | ...         |                |               |

Multiple blocks may be laid out horizontally (side by side) or vertically
(stacked with space between them) in the same sheet.

Category headers are matched against a comprehensive alias table that covers
both English names and Hebrew equivalents.
"""
from __future__ import annotations

import shutil
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

import openpyxl

from ledgerflow.config import Config
from ledgerflow.models import Category, Transaction

FIRST_DATA_ROW_INDEX = 6  # Excel row 7, 0-based for Python/openpyxl translation


# ------------------------------------------------------------------ #
# Exceptions                                                           #
# ------------------------------------------------------------------ #


class WorkbookStructureError(ValueError):
    """Base for all workbook-layout problems."""


class MonthSheetNotFoundError(WorkbookStructureError):
    """No sheet in the workbook matches the target month."""


class CategoryBlockNotFoundError(WorkbookStructureError):
    """No category block and no Other fallback exists for a transaction."""


# ------------------------------------------------------------------ #
# Month sheet matching constants                                       #
# ------------------------------------------------------------------ #

_HEBREW_MONTH_NAMES: dict[int, str] = {
    1:  "ינואר",
    2:  "פברואר",
    3:  "מרץ",
    4:  "אפריל",
    5:  "מאי",
    6:  "יוני",
    7:  "יולי",
    8:  "אוגוסט",
    9:  "ספטמבר",
    10: "אוקטובר",
    11: "נובמבר",
    12: "דצמבר",
}

_ENGLISH_MONTH_NAMES: dict[int, str] = {
    1:  "January",  2:  "February", 3:  "March",
    4:  "April",    5:  "May",      6:  "June",
    7:  "July",     8:  "August",   9:  "September",
    10: "October",  11: "November", 12: "December",
}

# Per-month ordered list of patterns tried for exact matching (case-insensitive).
_MONTH_PATTERNS: dict[int, list[str]] = {
    m: [
        f"{m:02d}",                          # "01" … "12"
        str(m),                              # "1" … "12"
        _HEBREW_MONTH_NAMES[m],
        _ENGLISH_MONTH_NAMES[m],             # "January" …
        _ENGLISH_MONTH_NAMES[m][:3],         # "Jan" …
    ]
    for m in range(1, 13)
}


# ------------------------------------------------------------------ #
# Category block detection constants                                   #
# ------------------------------------------------------------------ #

# Maps the lowercased cell value → Category.
# Covers English enum values, common Hebrew equivalents, and short aliases.
CATEGORY_ALIASES: dict[str, Category] = {
    # Other
    "other":    Category.OTHER,
    "אחר":      Category.OTHER,
    "שונות":    Category.OTHER,
    "כללי":     Category.OTHER,
    # Business
    "business": Category.BUSINESS,
    "עסקים":    Category.BUSINESS,
    "עסקי":     Category.BUSINESS,
    "עסק":      Category.BUSINESS,
    # Pharm
    "pharm":    Category.PHARM,
    "pharmacy": Category.PHARM,
    "פארמה":    Category.PHARM,
    "פרמה":     Category.PHARM,
    "רפואה":    Category.PHARM,
    "בריאות":   Category.PHARM,
    # Insurance
    "insurance":    Category.INSURANCE,
    "ביטוח":        Category.INSURANCE,
    # Transportation
    "transportation": Category.TRANSPORTATION,
    "תחבורה":       Category.TRANSPORTATION,
    "רכב":          Category.TRANSPORTATION,
    # Leisure
    "leisure":  Category.LEISURE,
    "פנאי":     Category.LEISURE,
    "בילוי":    Category.LEISURE,
    "בידור":    Category.LEISURE,
    # Home
    "home":     Category.HOME,
    "בית":      Category.HOME,
    "דירה":     Category.HOME,
    # Mortgage
    "mortgage":  Category.MORTGAGE,
    "משכנתא":   Category.MORTGAGE,
}


# Header labels used to infer the actual merchant|amount|date triplet
# when the category anchor is not the leftmost cell in the block.
MERCHANT_HEADER_ALIASES: set[str] = {
    "merchant",
    "description",
    "name",
    "details",
    "transaction description",
    "פירוט עסקה",
    "תיאור",
}

PRICE_HEADER_ALIASES: set[str] = {
    "price",
    "amount",
    "charge",
    "debit",
    "מחיר",
    "סכום",
    "סכום חיוב",
}

DATE_HEADER_ALIASES: set[str] = {
    "date",
    "purchase date",
    "transaction date",
    "charge date",
    "תאריך",
    "תאריך עסקה",
    "תאריך חיוב",
}


# ------------------------------------------------------------------ #
# Data types                                                           #
# ------------------------------------------------------------------ #


@dataclass(frozen=True)
class CategoryBlock:
    """
    Location of a single category block inside a sheet.

    All indices are **0-based** (matching Python list indexing).
    Add 1 when passing to openpyxl's ``ws.cell(row=..., column=...)``.
    """

    category: Category
    header_row: int    # row containing the category-name cell
    first_data_row: int  # first writable row below the block header area
    merchant_col: int  # column of the category/merchant name
    price_col: int     # column immediately to the right → amounts
    date_col: int      # column two to the right → dates


@dataclass(frozen=True)
class WriteResult:
    """Outcome of a single :func:`write_yearly_workbook` call."""

    backup_path: Path           # timestamped copy made before any write
    output_path: Path           # workbook written
    transactions_written: int
    skipped_count: int          # reserved; always 0 in the current implementation


# ------------------------------------------------------------------ #
# Public API                                                           #
# ------------------------------------------------------------------ #


def write_yearly_workbook(
    transactions: list[Transaction],
    source_workbook: Path,
    output_path: Path,
    month: int,
    config: Config,
    backup_dir: Optional[Path] = None,
) -> WriteResult:
    """
    Append *transactions* into the correct month sheet of *source_workbook*
    and save the result to *output_path*.

    The function never modifies *source_workbook* directly; it creates a
    timestamped backup first, then loads and modifies an in-memory copy,
    and saves only to *output_path*.

    Args:
        transactions:    Final categorized transactions (all from *month*).
        source_workbook: Existing yearly workbook that acts as the template.
        output_path:     Destination for the modified workbook.  Must differ
                         from *source_workbook*.
        month:           Calendar month (1–12) whose sheet will be modified.
        config:          Drives the allowed-category set.
        backup_dir:      Directory for the backup file.  Defaults to the
                         same directory as *source_workbook*.

    Returns:
        :class:`WriteResult` with backup path, output path, and row counts.

    Raises:
        ValueError:                 if *output_path* equals *source_workbook*.
        MonthSheetNotFoundError:    if no sheet matches *month*.
        CategoryBlockNotFoundError: if a transaction has no matching block
                                    and no Other block exists.
        OSError:                    if backup creation or workbook save fails.
    """
    if output_path.resolve() == source_workbook.resolve():
        raise ValueError(
            "output_path must differ from source_workbook — "
            "overwriting the source is not allowed."
        )

    # ── 1. Backup (must succeed before any other I/O) ─────────────────
    backup_path = create_backup(source_workbook, backup_dir)

    # ── 2. Load source workbook into memory ───────────────────────────
    wb = openpyxl.load_workbook(str(source_workbook))

    # ── 3. Locate the target month sheet ─────────────────────────────
    sheet_name = resolve_month_sheet(list(wb.sheetnames), month)
    ws = wb[sheet_name]

    # ── 4. Detect category blocks ─────────────────────────────────────
    rows: list[tuple] = [tuple(row) for row in ws.iter_rows(values_only=True)]
    blocks = detect_category_blocks(rows, config)

    # ── 5. Route each transaction to its block ─────────────────────────
    # Transactions with no dedicated block fall back to the Other block.
    # If neither the dedicated nor the Other block exists → raise.
    block_assignments: dict[CategoryBlock, list[Transaction]] = defaultdict(list)

    for tx in transactions:
        cat = tx.category if tx.category is not None else Category.OTHER
        if cat in blocks:
            block_assignments[blocks[cat]].append(tx)
        elif Category.OTHER in blocks:
            block_assignments[blocks[Category.OTHER]].append(tx)
        else:
            raise CategoryBlockNotFoundError(
                f"Transaction category {cat.value!r} has no block in sheet "
                f"{sheet_name!r}, and no 'Other' fallback block was detected. "
                f"Detected blocks: {sorted(b.value for b in blocks)}."
            )

    # ── 6. Write each block's transactions ────────────────────────────
    written = 0
    for block, block_txs in block_assignments.items():
        # Re-read the current sheet state so each block's search reflects
        # any rows already written in this run.
        current_rows: list[tuple] = [
            tuple(row) for row in ws.iter_rows(values_only=True)
        ]
        insert_row = find_first_empty_row(current_rows, block)

        for tx in block_txs:
            # openpyxl uses 1-based row/column indices.
            r = insert_row + 1
            ws.cell(row=r, column=block.merchant_col + 1, value=tx.merchant)
            price_cell = ws.cell(row=r, column=block.price_col + 1, value=float(tx.amount))
            price_cell.number_format = "#,##0.00"
            date_cell = ws.cell(row=r, column=block.date_col + 1, value=tx.date)
            date_cell.number_format = "dd/mm/yyyy"
            insert_row += 1
            written += 1

    # ── 7. Save to output_path (source is untouched) ──────────────────
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(output_path))

    return WriteResult(
        backup_path=backup_path,
        output_path=output_path,
        transactions_written=written,
        skipped_count=0,
    )


# ------------------------------------------------------------------ #
# Helpers — all exposed for direct unit testing                        #
# ------------------------------------------------------------------ #


def create_backup(source: Path, backup_dir: Optional[Path] = None) -> Path:
    """
    Copy *source* to a timestamped path and return that path.

    The timestamp format is ``YYYYMMDD_HHMMSS`` so backups sort
    chronologically.

    Raises ``OSError`` if the copy fails (disk full, permissions, etc.).
    """
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest_dir = backup_dir if backup_dir is not None else source.parent
    dest_dir.mkdir(parents=True, exist_ok=True)
    backup_path = dest_dir / f"{source.stem}_{ts}{source.suffix}"
    shutil.copy2(str(source), str(backup_path))
    return backup_path


def resolve_month_sheet(sheet_names: list[str], month: int) -> str:
    """
    Return the sheet name from *sheet_names* that best matches *month*.

    Matching order:
    1. Exact match (case-insensitive) against zero-padded number, bare
       number, Hebrew month name, English month name, 3-letter English.
    2. Contains-match (case-insensitive) as a fallback.

    ``wb.active`` is never returned implicitly — if nothing matches,
    :class:`MonthSheetNotFoundError` is raised.

    Args:
        sheet_names: Names from ``wb.sheetnames``.
        month:       Target calendar month (1–12).

    Raises:
        MonthSheetNotFoundError: if no sheet matches.
    """
    if not 1 <= month <= 12:
        raise ValueError(f"month must be 1–12, got {month!r}.")

    patterns = _MONTH_PATTERNS[month]
    patterns_lower = [p.lower() for p in patterns]

    # Pass 1 — exact match
    for sheet_name in sheet_names:
        if sheet_name.strip().lower() in patterns_lower:
            return sheet_name

    # Pass 2 — contains fallback
    for sheet_name in sheet_names:
        name_lower = sheet_name.strip().lower()
        for pattern in patterns_lower:
            if pattern and pattern in name_lower:
                return sheet_name

    raise MonthSheetNotFoundError(
        f"No sheet found for month {month} in the workbook. "
        f"Available sheets: {sheet_names}. "
        f"Expected patterns (any exact or contains match): {patterns}."
    )


def detect_category_blocks(
    rows: list[tuple],
    config: Config,
) -> dict[Category, CategoryBlock]:
    """
    Scan *rows* for category block headers and return a mapping of
    :class:`Category` → :class:`CategoryBlock`.

    A category block header is a cell whose value (stripped, lowercased)
    appears in :data:`CATEGORY_ALIASES` and whose category is in
    ``config.allowed_categories()``.

    The three columns of each block are:
    - ``col``     — merchant / category name
    - ``col + 1`` — price / amount
    - ``col + 2`` — date

    Only the **first** occurrence of each category is kept (left-to-right,
    top-to-bottom).

    Args:
        rows:   Sheet rows as plain Python tuples (e.g. from
                ``ws.iter_rows(values_only=True)``).
        config: Provides the allowed-category set.

    Returns:
        Dict mapping each detected :class:`Category` to its block descriptor.
    """
    allowed = set(config.allowed_categories())
    result: dict[Category, CategoryBlock] = {}

    for row_idx, row in enumerate(rows):
        for col_idx, cell in enumerate(row):
            if cell is None:
                continue
            canon = str(cell).strip().lower()
            category = CATEGORY_ALIASES.get(canon)
            if category is not None and category in allowed and category not in result:
                merchant_col, price_col, date_col, first_data_row = _resolve_block_layout(
                    rows, row_idx, col_idx
                )
                result[category] = CategoryBlock(
                    category=category,
                    header_row=row_idx,
                    first_data_row=max(first_data_row, FIRST_DATA_ROW_INDEX),
                    merchant_col=merchant_col,
                    price_col=price_col,
                    date_col=date_col,
                )

    return result


def find_first_empty_row(rows: list[tuple], block: CategoryBlock) -> int:
    """
    Return the **0-based** index of the first fully empty row below
    ``block.header_row``.

    A row is "fully empty" when all three of the block's target cells
    (merchant, price, date) are ``None`` or a blank/whitespace string.

    If no empty row exists within *rows*, the index one past the last row
    is returned so the caller can append new content.

    Args:
        rows:  Sheet rows as plain Python tuples.
        block: The category block whose insert point is sought.
    """
    mc, pc, dc = block.merchant_col, block.price_col, block.date_col

    def _empty(row: tuple, col: int) -> bool:
        if col >= len(row):
            return True
        v = row[col]
        return v is None or (isinstance(v, str) and not v.strip())

    for row_idx in range(max(block.first_data_row, FIRST_DATA_ROW_INDEX), len(rows)):
        row = rows[row_idx]
        # A row is writable when its merchant cell is absent.  This intentionally
        # ignores placeholder values (e.g. "1") that templates leave in the price
        # or date columns: those cells are overwritten when a real transaction is
        # written.  Only a real merchant name marks a row as already occupied.
        if _empty(row, mc):
            return row_idx

    return max(len(rows), FIRST_DATA_ROW_INDEX)  # append after the last existing row


def _resolve_block_layout(
    rows: list[tuple],
    row_idx: int,
    category_col: int,
) -> tuple[int, int, int, int]:
    # When the category label is in the leftmost column there is no room to
    # shift left, so fall straight to the simple layout.
    if category_col == 0:
        return (0, 1, 2, FIRST_DATA_ROW_INDEX)

    for candidate_row_idx in (row_idx, row_idx + 1):
        if candidate_row_idx >= len(rows):
            continue
        row = rows[candidate_row_idx]
        # Prefer a writable triplet to the right of the category label.
        for merchant_col in range(category_col + 1, category_col + 4):
            price_col = merchant_col + 1
            date_col = merchant_col + 2
            if _is_header_triplet(row, merchant_col, price_col, date_col):
                return (
                    max(0, merchant_col - 1),
                    max(0, price_col - 1),
                    max(0, date_col - 1),
                    FIRST_DATA_ROW_INDEX,
                )

        # Fallback: look for a triplet that starts to the left of the category.
        start_col = max(0, category_col - 2)
        end_col = category_col - 1
        for merchant_col in range(start_col, end_col + 1):
            price_col = merchant_col + 1
            date_col = merchant_col + 2
            if _is_header_triplet(row, merchant_col, price_col, date_col):
                return (
                    max(0, merchant_col - 1),
                    max(0, price_col - 1),
                    max(0, date_col - 1),
                    FIRST_DATA_ROW_INDEX,
                )

    # Default: the category label sits in the price column; merchant is one
    # column to the left, date one column to the right.
    return (
        category_col - 1,
        category_col,
        category_col + 1,
        FIRST_DATA_ROW_INDEX,
    )


def _is_header_triplet(
    row: tuple,
    merchant_col: int,
    price_col: int,
    date_col: int,
) -> bool:
    merchant_header = _cell_text(row, merchant_col)
    price_header = _cell_text(row, price_col)
    date_header = _cell_text(row, date_col)

    merchant_ok = merchant_header in MERCHANT_HEADER_ALIASES or merchant_header == ""
    return (
        merchant_ok
        and price_header in PRICE_HEADER_ALIASES
        and date_header in DATE_HEADER_ALIASES
    )


def _cell_text(row: tuple, col: int) -> str:
    if col >= len(row):
        return ""
    value = row[col]
    return str(value).strip().lower() if value is not None else ""
