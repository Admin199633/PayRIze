"""
Core domain models for LedgerFlow.

All entities are plain dataclasses with no framework dependencies.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from enum import Enum
from pathlib import Path
from typing import Optional


class Category(str, Enum):
    OTHER = "Other"
    BUSINESS = "Business"
    PHARM = "Pharm"
    INSURANCE = "Insurance"
    TRANSPORTATION = "Transportation"
    LEISURE = "Leisure"
    HOME = "Home"
    MORTGAGE = "Mortgage"  # optional; only used when config.include_mortgage is True


class ClassificationSource(str, Enum):
    SEED = "seed"
    CACHE = "cache"
    GPT = "gpt"
    FALLBACK = "fallback"
    RULE = "rule"  # deterministic post-processing rule (e.g. fuel convenience threshold)


# Dedup key: (date, normalized_merchant, rounded_amount)
# Defined as a type alias for readability; the actual value is a plain tuple.
DedupKey = tuple[date, str, Decimal]


@dataclass
class Transaction:
    """
    A single bank transaction extracted from a source file.

    - `merchant` holds the raw string from the spreadsheet.
    - `normalized_merchant` is the cleaned version used for dedup and categorization.
    - `category` is None until the categorization step runs.
    """

    date: date
    merchant: str
    normalized_merchant: str
    amount: Decimal
    source_file: Path
    sheet_name: str
    category: Optional[Category] = field(default=None)
    classification_source: Optional[ClassificationSource] = field(default=None)

    def __repr__(self) -> str:
        cat = self.category.value if self.category else "uncategorized"
        return (
            f"Transaction({self.date} | {self.merchant!r} | "
            f"{self.amount} | {cat})"
        )
