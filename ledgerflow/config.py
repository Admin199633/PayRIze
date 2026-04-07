"""
Configuration for LedgerFlow.

Loads from a JSON file or uses defaults. All paths remain as Path objects
so callers can resolve them relative to wherever they like.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from decimal import Decimal
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
    seed_spreadsheet_id: Optional[str] = None
    seed_worksheet_name: str = "Seed"
    cache_path: Optional[Path] = None
    year_spreadsheet_id: Optional[str] = None
    year_worksheet_names: dict[int, str] = field(default_factory=dict)
    run_log_worksheet_name: str = "RunLog"
    google_service_account_file: Optional[Path] = None
    currency_rates: dict[str, Decimal] = field(default_factory=dict)
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
        base_dir = path.parent

        cfg = cls()

        if "include_mortgage" in data:
            cfg.include_mortgage = bool(data["include_mortgage"])

        if cfg.include_mortgage and Category.MORTGAGE not in cfg.categories:
            cfg.categories.append(Category.MORTGAGE)

        if "seed_mapping_path" in data:
            cfg.seed_mapping_path = _resolve_config_path(data["seed_mapping_path"], base_dir)

        seed_cfg = data.get("seed", {})
        if isinstance(seed_cfg, dict):
            if "spreadsheet_id" in seed_cfg:
                cfg.seed_spreadsheet_id = str(seed_cfg["spreadsheet_id"])
            if "worksheet_name" in seed_cfg:
                cfg.seed_worksheet_name = str(seed_cfg["worksheet_name"])

        if "seed_spreadsheet_id" in data:
            cfg.seed_spreadsheet_id = str(data["seed_spreadsheet_id"])

        if "seed_worksheet_name" in data:
            cfg.seed_worksheet_name = str(data["seed_worksheet_name"])

        if "cache_path" in data:
            cfg.cache_path = _resolve_config_path(data["cache_path"], base_dir)

        if "year_spreadsheet_id" in data:
            cfg.year_spreadsheet_id = str(data["year_spreadsheet_id"])

        year_cfg = data.get("year", {})
        if isinstance(year_cfg, dict):
            if "spreadsheet_id" in year_cfg:
                cfg.year_spreadsheet_id = str(year_cfg["spreadsheet_id"])
            if "month_sheets" in year_cfg:
                cfg.year_worksheet_names = {
                    int(month): str(name)
                    for month, name in year_cfg["month_sheets"].items()
                }

        if "year_worksheet_names" in data:
            cfg.year_worksheet_names = {
                int(month): str(name)
                for month, name in data["year_worksheet_names"].items()
            }

        if "run_log_worksheet_name" in data:
            cfg.run_log_worksheet_name = str(data["run_log_worksheet_name"])

        if "google_service_account_file" in data:
            cfg.google_service_account_file = _resolve_config_path(
                data["google_service_account_file"],
                base_dir,
            )

        if "currency_rates" in data and isinstance(data["currency_rates"], dict):
            cfg.currency_rates = {
                str(code).upper(): Decimal(str(rate))
                for code, rate in data["currency_rates"].items()
            }

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

    def resolved_seed_target(self) -> tuple[str, str]:
        if not self.seed_spreadsheet_id:
            raise ValueError("seed.spreadsheet_id is not configured.")
        if not self.seed_worksheet_name:
            raise ValueError("seed.worksheet_name is not configured.")
        return self.seed_spreadsheet_id, self.seed_worksheet_name

    def resolved_year_target(self, month: int) -> tuple[str, str]:
        if not self.year_spreadsheet_id:
            raise ValueError("year.spreadsheet_id is not configured.")
        worksheet_name = self.year_worksheet_names.get(month)
        if not worksheet_name:
            raise ValueError(
                f"year.month_sheets does not contain an entry for month {month}."
            )
        return self.year_spreadsheet_id, worksheet_name

    def to_json_dict(self) -> dict:
        data: dict[str, object] = {}
        if self.include_mortgage:
            data["include_mortgage"] = True
        if self.seed_mapping_path is not None:
            data["seed_mapping_path"] = str(self.seed_mapping_path)
        if self.seed_spreadsheet_id or self.seed_worksheet_name != "Seed":
            data["seed"] = {
                "spreadsheet_id": self.seed_spreadsheet_id or "",
                "worksheet_name": self.seed_worksheet_name,
            }
        if self.cache_path is not None:
            data["cache_path"] = str(self.cache_path)
        if self.year_spreadsheet_id or self.year_worksheet_names:
            data["year"] = {
                "spreadsheet_id": self.year_spreadsheet_id or "",
                "month_sheets": {
                    str(month): name for month, name in self.year_worksheet_names.items()
                },
            }
        if self.run_log_worksheet_name != "RunLog":
            data["run_log_worksheet_name"] = self.run_log_worksheet_name
        if self.google_service_account_file is not None:
            data["google_service_account_file"] = str(self.google_service_account_file)
        if self.currency_rates:
            data["currency_rates"] = {
                code: float(rate) for code, rate in self.currency_rates.items()
            }
        if self.dedup_amount_rounding != 0:
            data["dedup_amount_rounding"] = self.dedup_amount_rounding
        if self.header_aliases != {k: list(v) for k, v in DEFAULT_HEADER_ALIASES.items()}:
            data["header_aliases"] = self.header_aliases
        if self.stop_tokens != DEFAULT_STOP_TOKENS.copy():
            data["stop_tokens"] = self.stop_tokens
        return data


def _resolve_config_path(raw_path: object, base_dir: Path) -> Path:
    resolved = Path(str(raw_path))
    if resolved.is_absolute():
        return resolved
    return base_dir / resolved
