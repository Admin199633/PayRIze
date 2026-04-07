"""
Amount parsing and rounding utilities.

Handles the variety of numeric representations found in bank exports:
currency symbols, thousands separators, parenthesised negatives, etc.
"""
from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Optional


# Regex that strips every character that is not a digit, dot, minus, or comma.
_NON_NUMERIC = re.compile(r"[^\d.\-,]")


def parse_amount(value: object) -> Optional[Decimal]:
    """
    Parse *value* into a :class:`~decimal.Decimal`.

    Handled cases:
    - int / float / Decimal  → converted directly
    - ``"$1,234.56"``        → 1234.56
    - ``"(50.00)"``          → -50.00  (accounting negative)
    - ``"-99.9"``            → -99.9
    - empty string / None    → None
    """
    if value is None:
        return None

    if isinstance(value, Decimal):
        return value

    if isinstance(value, (int, float)):
        return Decimal(str(value))

    if isinstance(value, str):
        raw = value.strip()
        if not raw:
            return None

        # Accounting-style negatives: (123.45) → -123.45
        negative = raw.startswith("(") and raw.endswith(")")
        if negative:
            raw = raw[1:-1]

        # Strip currency symbols, spaces, etc.
        raw = _NON_NUMERIC.sub("", raw)

        # Thousands commas: remove all commas (decimal point assumed to be ".")
        raw = raw.replace(",", "")

        if not raw or raw in ("-", "."):
            return None

        try:
            amount = Decimal(raw)
            return -amount if negative else amount
        except InvalidOperation:
            return None

    return None


def round_for_dedup(amount: Decimal, decimal_places: int = 0) -> Decimal:
    """
    Round *amount* to *decimal_places* for use in deduplication keys.

    Using ROUND_HALF_UP so that 0.5 rounds away from zero, consistent with
    common financial conventions.
    """
    quantize_str = "1" if decimal_places == 0 else "0." + "0" * decimal_places
    return amount.quantize(Decimal(quantize_str), rounding=ROUND_HALF_UP)
