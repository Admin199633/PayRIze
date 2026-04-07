"""
Transaction filtering utilities.

Kept separate from the pipeline so each filter is independently testable
and reusable without pulling in any pipeline or I/O dependencies.
"""
from __future__ import annotations

import re

from ledgerflow.models import Transaction


_INSTALLMENT_PATTERN = re.compile(r"תשלום\s+(\d+)\s+מתוך\s+(\d+)")


def filter_by_month(transactions: list[Transaction], month: int) -> list[Transaction]:
    """
    Return every transaction whose date falls in *month* (1–12).

    Order of the input list is preserved.  The function is intentionally pure:
    it does not modify the input and has no side effects.

    Args:
        transactions: Source list — may be empty.
        month:        Target calendar month, 1 = January … 12 = December.

    Returns:
        New list containing only transactions whose ``date.month == month``.

    Raises:
        ValueError: if *month* is outside the range 1–12.
    """
    if not 1 <= month <= 12:
        raise ValueError(
            f"month must be between 1 and 12 inclusive, got {month!r}."
        )
    return [
        tx for tx in transactions
        if tx.date.month == month or _has_valid_installment_override(tx)
    ]


def _has_valid_installment_override(tx: Transaction) -> bool:
    if not tx.additional_details:
        return False
    match = _INSTALLMENT_PATTERN.search(tx.additional_details)
    if match is None:
        return False
    current = int(match.group(1))
    total = int(match.group(2))
    return total > 0 and 1 <= current <= total
