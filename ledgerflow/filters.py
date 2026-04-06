"""
Transaction filtering utilities.

Kept separate from the pipeline so each filter is independently testable
and reusable without pulling in any pipeline or I/O dependencies.
"""
from __future__ import annotations

from ledgerflow.models import Transaction


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
    return [tx for tx in transactions if tx.date.month == month]
