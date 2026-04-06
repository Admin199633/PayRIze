"""
Excel extraction — reads bank export .xlsx/.xlsm files and produces Transaction objects.

Architecture
------------
openpyxl plumbing  (file → sheet → raw row lists of plain Python values)
    └─► _extract_from_sheet  (pure, testable — state-machine over raw rows)
            ├── detect_table_header   — (row) → column positions | None
            ├── is_stop_row           — recognises summary/total rows
            ├── is_all_empty_row      — recognises blank section separators
            └── parse_transaction_row — one raw row → Transaction | None

The helpers are intentionally not prefixed with _ so tests can import and
exercise them directly without going through a real workbook file.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import openpyxl

from ledgerflow.config import Config
from ledgerflow.models import Transaction
from ledgerflow.utils.amounts import parse_amount
from ledgerflow.utils.dates import parse_date
from ledgerflow.utils.normalization import canonicalize_header, normalize_merchant


# ------------------------------------------------------------------ #
# Exceptions                                                           #
# ------------------------------------------------------------------ #


class UnsupportedFileTypeError(ValueError):
    """Raised when the input file format is not supported."""


@dataclass(frozen=True)
class ExtractionStats:
    extracted_count: int = 0
    header_rows_detected: int = 0
    candidate_rows_seen: int = 0
    skipped_invalid_date: int = 0
    skipped_invalid_amount: int = 0
    skipped_missing_merchant: int = 0
    skipped_short_rows: int = 0

    def add(self, other: "ExtractionStats") -> "ExtractionStats":
        return ExtractionStats(
            extracted_count=self.extracted_count + other.extracted_count,
            header_rows_detected=self.header_rows_detected + other.header_rows_detected,
            candidate_rows_seen=self.candidate_rows_seen + other.candidate_rows_seen,
            skipped_invalid_date=self.skipped_invalid_date + other.skipped_invalid_date,
            skipped_invalid_amount=self.skipped_invalid_amount + other.skipped_invalid_amount,
            skipped_missing_merchant=self.skipped_missing_merchant + other.skipped_missing_merchant,
            skipped_short_rows=self.skipped_short_rows + other.skipped_short_rows,
        )


# ------------------------------------------------------------------ #
# Public API                                                           #
# ------------------------------------------------------------------ #


def extract_from_file(path: Path, config: Config) -> list[Transaction]:
    """
    Open *path* and extract all transactions from every sheet.

    Raises:
        UnsupportedFileTypeError: for .xls or any other unsupported extension.
        FileNotFoundError: propagated from openpyxl if the file is absent.
    """
    _validate_extension(path)

    return extract_from_file_with_stats(path, config)[0]


def extract_from_file_with_stats(
    path: Path,
    config: Config,
) -> tuple[list[Transaction], ExtractionStats]:
    wb = openpyxl.load_workbook(
        str(path),
        read_only=True,
        data_only=True,  # return cached values, not formula strings
    )
    try:
        result: list[Transaction] = []
        stats = ExtractionStats()
        for sheet in wb.worksheets:
            raw_rows = [list(row) for row in sheet.iter_rows(values_only=True)]
            sheet_transactions, sheet_stats = _extract_from_sheet_with_stats(
                raw_rows, path, sheet.title, config
            )
            result.extend(sheet_transactions)
            stats = stats.add(sheet_stats)
        return result, stats
    finally:
        wb.close()


def extract_from_files(paths: list[Path], config: Config) -> list[Transaction]:
    """
    Extract and concatenate transactions from multiple Excel files.

    Files are processed independently; deduplication is the caller's responsibility.
    """
    return extract_from_files_with_stats(paths, config)[0]


def extract_from_files_with_stats(
    paths: list[Path],
    config: Config,
) -> tuple[list[Transaction], ExtractionStats]:
    result: list[Transaction] = []
    stats = ExtractionStats()
    for path in paths:
        file_transactions, file_stats = extract_from_file_with_stats(path, config)
        result.extend(file_transactions)
        stats = stats.add(file_stats)
    return result, stats


# ------------------------------------------------------------------ #
# Sheet-level extraction  (pure — no openpyxl types)                  #
# ------------------------------------------------------------------ #


def _extract_from_sheet(
    rows: list[list[object]],
    source_file: Path,
    sheet_name: str,
    config: Config,
) -> list[Transaction]:
    return _extract_from_sheet_with_stats(rows, source_file, sheet_name, config)[0]


def _extract_from_sheet_with_stats(
    rows: list[list[object]],
    source_file: Path,
    sheet_name: str,
    config: Config,
) -> tuple[list[Transaction], ExtractionStats]:
    """
    Run a state-machine over *rows* and collect every parseable transaction.

    For each row, in order:
      1. If it looks like a header row → record column positions, enter table mode.
         (This also handles repeated / restarted tables on the same sheet.)
      2. If currently in table mode:
         a. Blank row or stop-token row → exit table mode.
         b. Otherwise → attempt to parse a Transaction; append if successful.
      3. Rows before the first header are silently ignored.
    """
    transactions: list[Transaction] = []
    stats = ExtractionStats()
    in_table = False
    date_col = merchant_col = amount_col = -1

    for row in rows:
        header = detect_table_header(row, config)
        if header is not None:
            date_col, merchant_col, amount_col = header
            in_table = True
            stats = stats.add(ExtractionStats(header_rows_detected=1))
            continue

        if not in_table:
            continue

        if is_all_empty_row(row) or is_stop_row(row, config):
            in_table = False
            date_col = merchant_col = amount_col = -1
            continue

        stats = stats.add(ExtractionStats(candidate_rows_seen=1))
        tx, reason = _parse_transaction_row_with_reason(
            row, date_col, merchant_col, amount_col, source_file, sheet_name
        )
        if tx is not None:
            transactions.append(tx)
            stats = stats.add(ExtractionStats(extracted_count=1))
        elif reason == "short_row":
            stats = stats.add(ExtractionStats(skipped_short_rows=1))
        elif reason == "invalid_date":
            stats = stats.add(ExtractionStats(skipped_invalid_date=1))
        elif reason == "missing_merchant":
            stats = stats.add(ExtractionStats(skipped_missing_merchant=1))
        elif reason == "invalid_amount":
            stats = stats.add(ExtractionStats(skipped_invalid_amount=1))

    return transactions, stats


# ------------------------------------------------------------------ #
# Helpers — exposed for unit testing                                   #
# ------------------------------------------------------------------ #


def detect_table_header(
    row: list[object],
    config: Config,
) -> Optional[tuple[int, int, int]]:
    """
    If *row* is a valid transaction-table header, return ``(date_col, merchant_col, amount_col)``.

    Matching is case-insensitive and uses the configured aliases.  Each logical
    column is assigned to the first cell (left-to-right) whose canonicalised
    value appears in the corresponding alias set.

    Returns ``None`` if all three columns are not found.
    """
    date_aliases = config.header_aliases_flat("purchase_date")
    merchant_aliases = config.header_aliases_flat("merchant")
    amount_aliases = config.header_aliases_flat("amount")

    date_col: Optional[int] = None
    merchant_col: Optional[int] = None
    amount_col: Optional[int] = None

    for col_idx, cell in enumerate(row):
        if cell is None:
            continue
        canon = canonicalize_header(str(cell))
        if not canon:
            continue
        if date_col is None and _matches_header_alias(canon, date_aliases):
            date_col = col_idx
        elif merchant_col is None and _matches_header_alias(canon, merchant_aliases):
            merchant_col = col_idx
        elif amount_col is None and _matches_header_alias(canon, amount_aliases):
            amount_col = col_idx

        if date_col is not None and merchant_col is not None and amount_col is not None:
            break  # early exit once all three are found

    if date_col is None or merchant_col is None or amount_col is None:
        return None

    return (date_col, merchant_col, amount_col)


def is_stop_row(cells: list[object], config: Config) -> bool:
    """
    Return ``True`` if any non-empty cell in *cells* contains a stop token
    (case-insensitive substring match).

    Stop tokens mark summary, total, and section-header rows that signal the
    end of a transaction table.

    Known limitation: a merchant whose name *contains* a stop token (e.g.
    "Total Sports") would also be treated as a stop row.  If this causes false
    positives, remove the conflicting token from ``config.stop_tokens``.
    """
    stop_tokens = [t.lower() for t in config.stop_tokens]
    for cell in cells:
        if cell is None:
            continue
        text = str(cell).strip().lower()
        if not text:
            continue
        for token in stop_tokens:
            if token in text:
                return True
    return False


def is_all_empty_row(cells: list[object]) -> bool:
    """Return ``True`` if every cell in *cells* is ``None`` or blank whitespace."""
    for cell in cells:
        if cell is None:
            continue
        if isinstance(cell, str) and not cell.strip():
            continue
        return False
    return True


def parse_transaction_row(
    cells: list[object],
    date_col: int,
    merchant_col: int,
    amount_col: int,
    source_file: Path,
    sheet_name: str,
) -> Optional[Transaction]:
    """
    Parse one data row into a :class:`~ledgerflow.models.Transaction`.

    Returns ``None`` (without raising) for rows that are missing a valid date,
    merchant name, or amount.  Invalid rows are silently skipped so a single
    bad row never aborts the whole extraction run.
    """
    tx, _ = _parse_transaction_row_with_reason(
        cells, date_col, merchant_col, amount_col, source_file, sheet_name
    )
    return tx


def _parse_transaction_row_with_reason(
    cells: list[object],
    date_col: int,
    merchant_col: int,
    amount_col: int,
    source_file: Path,
    sheet_name: str,
) -> tuple[Optional[Transaction], str | None]:
    max_col = max(date_col, merchant_col, amount_col)
    if len(cells) <= max_col:
        return None, "short_row"

    parsed_date = parse_date(cells[date_col])
    if parsed_date is None:
        return None, "invalid_date"

    raw_merchant = cells[merchant_col]
    merchant_str = str(raw_merchant).strip() if raw_merchant is not None else ""
    if not merchant_str:
        return None, "missing_merchant"

    parsed_amount = parse_amount(cells[amount_col])
    if parsed_amount is None:
        return None, "invalid_amount"

    return Transaction(
        date=parsed_date,
        merchant=merchant_str,
        normalized_merchant=normalize_merchant(merchant_str),
        amount=parsed_amount,
        source_file=source_file,
        sheet_name=sheet_name,
    ), None


# ------------------------------------------------------------------ #
# Private                                                              #
# ------------------------------------------------------------------ #

_SUPPORTED_EXTENSIONS: frozenset[str] = frozenset({".xlsx", ".xlsm"})
_XLS_EXTENSION = ".xls"


def _validate_extension(path: Path) -> None:
    suffix = path.suffix.lower()
    if suffix == _XLS_EXTENSION:
        raise UnsupportedFileTypeError(
            f"{path.name!r} is a legacy .xls file. "
            "Convert it to .xlsx first (File \u2192 Save As in Excel) \u2014 "
            "openpyxl does not support the older binary format."
        )
    if suffix not in _SUPPORTED_EXTENSIONS:
        raise UnsupportedFileTypeError(
            f"{path.name!r} has unsupported extension {suffix!r}. "
            f"Supported formats: {', '.join(sorted(_SUPPORTED_EXTENSIONS))}."
        )


def _matches_header_alias(canonical_value: str, aliases: set[str]) -> bool:
    if canonical_value in aliases:
        return True
    for alias in aliases:
        if not canonical_value.startswith(alias):
            continue
        if len(canonical_value) == len(alias):
            return True
        if canonical_value[len(alias)] in {" ", ":", "-", "_", "/"}:
            return True
    return False
