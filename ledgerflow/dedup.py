"""
Deduplication logic for merged transaction lists.

A duplicate is defined as two transactions that share the same:
  - date
  - normalized merchant name
  - amount (rounded to config-controlled precision)

The first occurrence (by list order) is always kept.
"""
from __future__ import annotations

from ledgerflow.models import DedupKey, Transaction
from ledgerflow.utils.amounts import round_for_dedup


def make_dedup_key(tx: Transaction, amount_decimal_places: int = 0) -> DedupKey:
    """
    Build the deduplication key for *tx*.

    The key is a plain tuple so it is hashable and can be stored in a set.
    """
    rounded = round_for_dedup(tx.amount, amount_decimal_places)
    return (tx.date, tx.normalized_merchant, rounded)


def deduplicate(
    transactions: list[Transaction],
    amount_decimal_places: int = 0,
) -> list[Transaction]:
    """
    Return a new list with duplicate transactions removed.

    First occurrence wins; relative order of surviving transactions is preserved.

    Args:
        transactions: Input list (may come from multiple source files).
        amount_decimal_places: Precision used when rounding amounts for the key.
    """
    seen: set[DedupKey] = set()
    result: list[Transaction] = []

    for tx in transactions:
        key = make_dedup_key(tx, amount_decimal_places)
        if key not in seen:
            seen.add(key)
            result.append(tx)

    return result


def find_duplicates(
    transactions: list[Transaction],
    amount_decimal_places: int = 0,
) -> dict[DedupKey, list[Transaction]]:
    """
    Group transactions by dedup key, returning only groups with more than one entry.

    Useful for inspecting which rows were merged and why.
    """
    groups: dict[DedupKey, list[Transaction]] = {}

    for tx in transactions:
        key = make_dedup_key(tx, amount_decimal_places)
        groups.setdefault(key, []).append(tx)

    return {key: group for key, group in groups.items() if len(group) > 1}
