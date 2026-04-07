"""
CSV export for transaction lists.

Two separate functions with explicit, stable column orders:
  - write_raw_csv        — filtered transactions before categorization
  - write_categorized_csv — final transactions after all categorization passes

Encoding: UTF-8 with BOM (utf-8-sig) so that Excel on Windows opens Hebrew
text correctly without a manual import wizard.
"""
from __future__ import annotations

import csv
from pathlib import Path

from ledgerflow.models import Transaction


# Column order is intentional and must not change between runs, so that
# downstream tools (yearly workbook writer, Excel macros) can rely on it.

_RAW_COLUMNS: tuple[str, ...] = (
    "date",
    "merchant",
    "normalized_merchant",
    "amount",
    "source_file",
    "sheet_name",
)

_CATEGORIZED_COLUMNS: tuple[str, ...] = (
    "date",
    "merchant",
    "normalized_merchant",
    "amount",
    "category",
    "classification_source",
    "source_file",
    "sheet_name",
)


def write_raw_csv(transactions: list[Transaction], path: Path) -> None:
    """
    Write *transactions* to *path* as a CSV with raw (pre-categorization) columns.

    Creates parent directories if they do not exist.
    Existing files are overwritten.

    Column order: date, merchant, normalized_merchant, amount, source_file, sheet_name.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        writer.writerow(_RAW_COLUMNS)
        for tx in transactions:
            writer.writerow(_raw_row(tx))


def write_categorized_csv(transactions: list[Transaction], path: Path) -> None:
    """
    Write *transactions* to *path* as a CSV with the final category column included.

    Creates parent directories if they do not exist.
    Existing files are overwritten.

    Column order: date, merchant, normalized_merchant, amount, category,
                  source_file, sheet_name.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        writer.writerow(_CATEGORIZED_COLUMNS)
        for tx in transactions:
            writer.writerow(_categorized_row(tx))


# ------------------------------------------------------------------ #
# Column accessors (exposed so tests can verify individual cell values)
# ------------------------------------------------------------------ #

RAW_COLUMNS = _RAW_COLUMNS
CATEGORIZED_COLUMNS = _CATEGORIZED_COLUMNS


def _raw_row(tx: Transaction) -> list[str]:
    return [
        tx.date.isoformat(),
        tx.merchant,
        tx.normalized_merchant,
        str(tx.amount),
        str(tx.source_file),
        tx.sheet_name,
    ]


def _categorized_row(tx: Transaction) -> list[str]:
    return [
        tx.date.isoformat(),
        tx.merchant,
        tx.normalized_merchant,
        str(tx.amount),
        tx.category.value if tx.category is not None else "",
        tx.classification_source.value if tx.classification_source is not None else "",
        str(tx.source_file),
        tx.sheet_name,
    ]
