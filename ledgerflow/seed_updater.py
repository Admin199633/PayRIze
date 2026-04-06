"""
Seed file updater.

Provides two capabilities:
1. Extract the GPT-classified transactions from a completed run so they can
   be reviewed by the user.
2. Write the reviewed merchant → category decisions back into the seed Excel
   file, growing it over time so GPT is needed less on future runs.

Seed file format (same as what ``categorization.load_seed_mapping`` reads):

    | MerchantKey | MerchantOriginal | Category | SourceFile | SourceSheet |
    | ...         | ...              | ...      | ...        | ...         |

``MerchantKey`` and ``MerchantOriginal`` are the only columns required by the
reader.  ``SourceFile`` and ``SourceSheet`` are optional metadata — this module
preserves them when the file already has those columns.
"""
from __future__ import annotations

from pathlib import Path
from typing import Optional

import openpyxl

from ledgerflow.models import Category, ClassificationSource, Transaction
from ledgerflow.utils.normalization import normalize_merchant

# ------------------------------------------------------------------ #
# Numeric category map (used by the review UI)                        #
# ------------------------------------------------------------------ #

NUMERIC_CATEGORY_MAP: dict[int, Category] = {
    1: Category.HOME,
    2: Category.LEISURE,
    3: Category.TRANSPORTATION,
    4: Category.INSURANCE,
    5: Category.PHARM,
    6: Category.BUSINESS,
    7: Category.MORTGAGE,
    8: Category.OTHER,
}

# Reverse map for display
_CATEGORY_TO_NUMBER: dict[Category, int] = {v: k for k, v in NUMERIC_CATEGORY_MAP.items()}

# Required seed column headers (case-insensitive canonical forms)
_REQUIRED_HEADERS = ("merchantkey", "merchantoriginal", "category")
_OPTIONAL_HEADERS = ("sourcefile", "sourcesheet")


# ------------------------------------------------------------------ #
# Public helpers                                                       #
# ------------------------------------------------------------------ #


def extract_gpt_transactions(transactions: list[Transaction]) -> list[Transaction]:
    """Return only those transactions whose classification source was GPT."""
    return [
        tx for tx in transactions
        if tx.classification_source == ClassificationSource.GPT
    ]


def category_number(category: Category) -> Optional[int]:
    """Return the numeric key for *category*, or None if unmapped."""
    return _CATEGORY_TO_NUMBER.get(category)


def update_seed_file(
    path: Path,
    reviewed: dict[str, tuple[str, Category]],
) -> None:
    """
    Append or update merchant → category entries in the seed Excel file.

    Args:
        path:     Path to the seed workbook (.xlsx / .xlsm).  The file is
                  created from scratch (with the standard header row) if it
                  does not yet exist.
        reviewed: Mapping of ``normalized_merchant →
                  (merchant_original, category)``.  Each entry is either
                  appended as a new row or overwrites the existing row for
                  that merchant key, whichever is found first.

    Behaviour
    ---------
    * **Deduplication**: if the same normalized merchant key already exists in
      the file, its row is updated in-place (category replaced).  The file
      will never contain two rows with the same ``MerchantKey``.
    * **Deterministic update**: the *first* occurrence of a duplicate key is
      updated; subsequent occurrences are left untouched (they represent old
      data that the user should clean up manually).
    * **Column preservation**: if the existing file has ``SourceFile`` and
      ``SourceSheet`` columns their positions are retained; new entries leave
      those cells blank.
    """
    if not reviewed:
        return

    if path.exists():
        _update_existing(path, reviewed)
    else:
        _create_new(path, reviewed)


# ------------------------------------------------------------------ #
# Private helpers                                                      #
# ------------------------------------------------------------------ #


def _update_existing(
    path: Path,
    reviewed: dict[str, tuple[str, Category]],
) -> None:
    wb = openpyxl.load_workbook(str(path))
    ws = wb.active

    # ── Locate the header row and column positions ────────────────────
    header_row_idx: Optional[int] = None
    col_map: dict[str, int] = {}  # lower-header → 1-based column index

    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        for col_idx, cell in enumerate(row, start=1):
            if cell is None:
                continue
            canon = str(cell).strip().lower()
            if canon in _REQUIRED_HEADERS or canon in _OPTIONAL_HEADERS:
                col_map[canon] = col_idx
        if all(h in col_map for h in _REQUIRED_HEADERS):
            header_row_idx = row_idx
            break

    if header_row_idx is None:
        # No valid header found — append a fresh header block at the end.
        _append_header_and_rows(ws, reviewed, after_row=ws.max_row)
        wb.save(str(path))
        return

    key_col = col_map["merchantkey"]
    original_col = col_map["merchantoriginal"]
    category_col = col_map["category"]
    source_file_col = col_map.get("sourcefile")
    source_sheet_col = col_map.get("sourcesheet")

    # ── Build set of already-seen merchant keys (first occurrence only) ─
    seen: dict[str, int] = {}  # normalized_key → row number (1-based)
    for row_num in range(header_row_idx + 1, ws.max_row + 1):
        cell_val = ws.cell(row=row_num, column=key_col).value
        if cell_val is None:
            continue
        norm = normalize_merchant(str(cell_val))
        if norm and norm not in seen:
            seen[norm] = row_num

    # ── Update existing rows / collect new entries ─────────────────────
    new_entries: dict[str, tuple[str, Category]] = {}
    for norm_key, (original, category) in reviewed.items():
        if norm_key in seen:
            row_num = seen[norm_key]
            ws.cell(row=row_num, column=key_col).value = norm_key
            ws.cell(row=row_num, column=original_col).value = original
            ws.cell(row=row_num, column=category_col).value = category.value
            # Source columns are intentionally left unchanged for existing rows.
        else:
            new_entries[norm_key] = (original, category)

    # ── Append brand-new entries after the last data row ──────────────
    insert_row = ws.max_row + 1
    for norm_key, (original, category) in new_entries.items():
        ws.cell(row=insert_row, column=key_col).value = norm_key
        ws.cell(row=insert_row, column=original_col).value = original
        ws.cell(row=insert_row, column=category_col).value = category.value
        # Leave optional columns blank for new rows.
        insert_row += 1

    wb.save(str(path))


def _create_new(
    path: Path,
    reviewed: dict[str, tuple[str, Category]],
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["MerchantKey", "MerchantOriginal", "Category"])
    _append_header_and_rows(ws, reviewed, after_row=1)
    wb.save(str(path))


def _append_header_and_rows(
    ws: object,
    reviewed: dict[str, tuple[str, Category]],
    after_row: int,
) -> None:
    """Append data rows directly (header already written or not needed here)."""
    # For _create_new the header is row 1; rows start at 2.
    # For _update_existing with no header, we add header + data.
    for norm_key, (original, category) in reviewed.items():
        ws.append([norm_key, original, category.value])
