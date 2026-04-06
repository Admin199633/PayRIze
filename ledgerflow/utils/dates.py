"""
Date parsing utilities.

Accepts datetime objects, date objects, strings in common formats, and
openpyxl-style Excel serial numbers (as int/float fallback).
"""
from __future__ import annotations

from datetime import date, datetime
from typing import Optional


# Ordered by likelihood of appearance in Israeli bank exports.
_DATE_FORMATS: list[str] = [
    "%d/%m/%Y",   # 15/03/2024  (most common in IL exports)
    "%Y-%m-%d",   # 2024-03-15  (ISO)
    "%d-%m-%Y",   # 15-03-2024
    "%m/%d/%Y",   # 03/15/2024  (US format)
    "%d.%m.%Y",   # 15.03.2024
    "%d.%m.%y",   # 15.03.24    (2-digit year with dots)
    "%Y/%m/%d",   # 2024/03/15
    "%d/%m/%y",   # 15/03/24   (2-digit year)
    "%Y%m%d",     # 20240315   (compact)
]


def parse_date(value: object) -> Optional[date]:
    """
    Parse *value* into a :class:`datetime.date`.
    Returns ``None`` if the value cannot be interpreted as a date.
    """
    if value is None:
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    if isinstance(value, (int, float)):
        # openpyxl converts most Excel dates automatically, but handle numeric
        # serials that slip through (e.g. cached formula values).
        try:
            from openpyxl.utils.datetime import from_excel  # type: ignore[import]
            return from_excel(int(value)).date()
        except Exception:
            return None

    if isinstance(value, str):
        cleaned = value.strip()
        for fmt in _DATE_FORMATS:
            try:
                return datetime.strptime(cleaned, fmt).date()
            except ValueError:
                continue

    return None


def is_valid_date(value: object) -> bool:
    """Return True if *value* can be parsed into a date."""
    return parse_date(value) is not None


def month_of(d: date) -> int:
    """Return the 1-based month number of *d*."""
    return d.month
