"""
Configuration for LedgerFlow.

Loads from a JSON file or uses defaults. All paths remain as Path objects
so callers can resolve them relative to wherever they like.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from ledgerflow.models import Category

# Default allowed categories (Mortgage excluded unless opted-in via config).
DEFAULT_CATEGORIES: list[Category] = [
    Category.OTHER,
    Category.BUSINESS,
    Category.PHARM,
    Category.INSURANCE,
    Category.TRANSPORTATION,
    Category.LEISURE,
    Category.HOME,
]

# Column-header synonyms (lowercase) for the three required columns.
# Add translated / locale-specific headers here as the project grows.
DEFAULT_HEADER_ALIASES: dict[str, list[str]] = {
    "purchase_date": [
        "purchase date",
        "transaction date",
        "charge date",
        "date",
        "תאריך עסקה",
        "תאריך",
    ],
    "merchant": [
        "merchant",
        "business name",
        "description",
        "שם בית עסק",
        "תיאור",
        "name",
    ],
    "amount": [
        "amount",
        "charge",
        "debit",
        "סכום חיוב",
        "סכום",
        "total",
    ],
}

# Sentinel values in cells that signal the end of a transaction table
# (summary rows, section headers, etc.).
DEFAULT_STOP_TOKENS: list[str] = [
    "total",
    "סה\"כ",
    "סהכ",
    "grand total",
    "subtotal",
]


@dataclass
class Config:
    categories: list[Category] = field(
        default_factory=lambda: DEFAULT_CATEGORIES.copy()
    )
    include_mortgage: bool = False
    seed_mapping_path: Optional[Path] = None
    cache_path: Optional[Path] = None
    # Decimal places used when rounding amounts for dedup (0 = whole number).
    dedup_amount_rounding: int = 0
    header_aliases: dict[str, list[str]] = field(
        default_factory=lambda: {k: list(v) for k, v in DEFAULT_HEADER_ALIASES.items()}
    )
    stop_tokens: list[str] = field(
        default_factory=lambda: DEFAULT_STOP_TOKENS.copy()
    )

    # ------------------------------------------------------------------ #
    # Factory                                                              #
    # ------------------------------------------------------------------ #

    @classmethod
    def from_json(cls, path: Path) -> Config:
        """Load config from a JSON file, falling back to defaults for missing keys."""
        with path.open("r", encoding="utf-8") as fh:
            data: dict = json.load(fh)

        cfg = cls()

        if "include_mortgage" in data:
            cfg.include_mortgage = bool(data["include_mortgage"])

        if cfg.include_mortgage and Category.MORTGAGE not in cfg.categories:
            cfg.categories.append(Category.MORTGAGE)

        if "seed_mapping_path" in data:
            cfg.seed_mapping_path = Path(data["seed_mapping_path"])

        if "cache_path" in data:
            cfg.cache_path = Path(data["cache_path"])

        if "dedup_amount_rounding" in data:
            cfg.dedup_amount_rounding = int(data["dedup_amount_rounding"])

        if "header_aliases" in data:
            # Merge caller overrides on top of defaults.
            for key, aliases in data["header_aliases"].items():
                cfg.header_aliases[key] = list(aliases)

        if "stop_tokens" in data:
            cfg.stop_tokens = list(data["stop_tokens"])

        return cfg

    # ------------------------------------------------------------------ #
    # Helpers                                                              #
    # ------------------------------------------------------------------ #

    def allowed_categories(self) -> list[Category]:
        """Return the categories permitted for this run."""
        if self.include_mortgage:
            return list(self.categories)
        return [c for c in self.categories if c != Category.MORTGAGE]

    def header_aliases_flat(self, column: str) -> set[str]:
        """
        Return the full set of lowercase aliases for a logical column name.
        ``column`` must be one of 'purchase_date', 'merchant', 'amount'.
        """
        return {alias.lower() for alias in self.header_aliases.get(column, [])}
