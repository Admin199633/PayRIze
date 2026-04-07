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
import re
from typing import Optional

import openpyxl

from ledgerflow.config import Config
from ledgerflow.models import Transaction
from ledgerflow.utils.amounts import parse_amount
from ledgerflow.utils.dates import parse_date
from ledgerflow.utils.exchange_rates import convert_amount_to_ils
from ledgerflow.utils.normalization import canonicalize_header, normalize_merchant


# ------------------------------------------------------------------ #
# Exceptions                                                           #
# ------------------------------------------------------------------ #


class UnsupportedFileTypeError(ValueError):
    """Raised when the input file format is not supported."""


_INSTALLMENT_PATTERN = re.compile(r"תשלום\s+(\d+)\s+מתוך\s+(\d+)")
_CURRENCY_ALIASES = {
    "transaction_currency": {
        "transaction currency",
        "currency",
        "מטבע עסקה",
    },
    "charge_currency": {
        "charge currency",
        "debit currency",
        "מטבע חיוב",
    },
    "charge_date": {
        "bank charge date",
        "account charge date",
        "חיוב בחשבון הבנק",
    },
}


@dataclass(frozen=True)
class _TableLayout:
    date_col: int
    merchant_col: int
    amount_col: int
    transaction_currency_col: int | None = None
    charge_currency_col: int | None = None
    charge_date_col: int | None = None


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
    layout: _TableLayout | None = None

    for row in rows:
        detected_layout = _detect_table_layout(row, config)
        if detected_layout is not None:
            layout = detected_layout
            in_table = True
            stats = stats.add(ExtractionStats(header_rows_detected=1))
            continue

        if not in_table:
            continue

        if is_all_empty_row(row) or is_stop_row(row, config):
            in_table = False
            layout = None
            continue

        assert layout is not None
        stats = stats.add(ExtractionStats(candidate_rows_seen=1))
        tx, reason = _parse_transaction_row_with_reason(
            row,
            layout.date_col,
            layout.merchant_col,
            layout.amount_col,
            source_file,
            sheet_name,
            config=config,
            transaction_currency_col=layout.transaction_currency_col,
            charge_currency_col=layout.charge_currency_col,
            charge_date_col=layout.charge_date_col,
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
    layout = _detect_table_layout(row, config)
    if layout is None:
        return None
    return (layout.date_col, layout.merchant_col, layout.amount_col)


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
        cells, date_col, merchant_col, amount_col, source_file, sheet_name, config=Config()
    )
    return tx


def _parse_transaction_row_with_reason(
    cells: list[object],
    date_col: int,
    merchant_col: int,
    amount_col: int,
    source_file: Path,
    sheet_name: str,
    config: Config,
    transaction_currency_col: int | None = None,
    charge_currency_col: int | None = None,
    charge_date_col: int | None = None,
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

    effective_currency = _extract_effective_currency(
        cells,
        transaction_currency_col=transaction_currency_col,
        charge_currency_col=charge_currency_col,
    )
    if effective_currency and not _is_ils_currency(effective_currency):
        conversion_date = _extract_conversion_date(
            cells,
            purchase_date=parsed_date,
            charge_date_col=charge_date_col,
        )
        parsed_amount = convert_amount_to_ils(
            parsed_amount,
            effective_currency,
            conversion_date,
            config,
        )

    return Transaction(
        date=parsed_date,
        merchant=merchant_str,
        normalized_merchant=normalize_merchant(merchant_str),
        amount=parsed_amount,
        source_file=source_file,
        sheet_name=sheet_name,
        additional_details=_extract_installment_details(cells),
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


def _select_effective_amount_column(
    amount_candidates: list[tuple[int, str]],
) -> Optional[int]:
    if not amount_candidates:
        return None

    for col_idx, canonical_header in amount_candidates:
        if _is_charge_amount_header(canonical_header):
            return col_idx

    return amount_candidates[0][0]


def _detect_table_layout(
    row: list[object],
    config: Config,
) -> Optional[_TableLayout]:
    date_aliases = config.header_aliases_flat("purchase_date")
    merchant_aliases = config.header_aliases_flat("merchant")
    amount_aliases = config.header_aliases_flat("amount")

    date_col: Optional[int] = None
    merchant_col: Optional[int] = None
    amount_candidates: list[tuple[int, str]] = []
    transaction_currency_col: int | None = None
    charge_currency_col: int | None = None
    charge_date_col: int | None = None

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
        elif _matches_header_alias(canon, amount_aliases):
            amount_candidates.append((col_idx, canon))
        elif transaction_currency_col is None and _matches_header_alias(
            canon, _CURRENCY_ALIASES["transaction_currency"]
        ):
            transaction_currency_col = col_idx
        elif charge_currency_col is None and _matches_header_alias(
            canon, _CURRENCY_ALIASES["charge_currency"]
        ):
            charge_currency_col = col_idx
        elif charge_date_col is None and _matches_header_alias(
            canon, _CURRENCY_ALIASES["charge_date"]
        ):
            charge_date_col = col_idx

    amount_col = _select_effective_amount_column(amount_candidates)
    if date_col is None or merchant_col is None or amount_col is None:
        return None
    return _TableLayout(
        date_col=date_col,
        merchant_col=merchant_col,
        amount_col=amount_col,
        transaction_currency_col=transaction_currency_col,
        charge_currency_col=charge_currency_col,
        charge_date_col=charge_date_col,
    )


def _is_charge_amount_header(canonical_header: str) -> bool:
    charge_aliases = (
        "charge",
        "charge amount",
        "debit",
        "debit amount",
        "סכום חיוב",
    )
    return any(
        canonical_header == alias or canonical_header.startswith(f"{alias} ")
        for alias in charge_aliases
    )


def _extract_installment_details(cells: list[object]) -> str:
    for cell in cells:
        if cell is None:
            continue
        text = str(cell).strip()
        if not text:
            continue
        if _INSTALLMENT_PATTERN.search(text):
            return text
    return ""


def _extract_effective_currency(
    cells: list[object],
    *,
    transaction_currency_col: int | None,
    charge_currency_col: int | None,
) -> str:
    if charge_currency_col is not None and charge_currency_col < len(cells):
        value = _normalize_currency(cells[charge_currency_col])
        if value:
            return value
    if transaction_currency_col is not None and transaction_currency_col < len(cells):
        value = _normalize_currency(cells[transaction_currency_col])
        if value:
            return value
    return ""


def _extract_conversion_date(
    cells: list[object],
    *,
    purchase_date,
    charge_date_col: int | None,
):
    if charge_date_col is not None and charge_date_col < len(cells):
        charge_date = parse_date(cells[charge_date_col])
        if charge_date is not None:
            return charge_date
    return purchase_date


def _normalize_currency(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip().upper()
    if text in {"₪", "ש\"ח", "שח", "ILS", "NIS"}:
        return "ILS"
    if text in {"$", "USD"}:
        return "USD"
    if text in {"€", "EUR"}:
        return "EUR"
    return text


def _is_ils_currency(currency_code: str) -> bool:
    return currency_code.upper() == "ILS"
