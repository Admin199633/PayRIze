"""
Transaction categorization.

Pipeline (executed in order, first valid match wins):
  1. Seed mapping  — static Excel file curated by the user; highest trust.
  2. Cache         — JSON file persisted between runs, populated by GPT step.
  3. Fallback      — Category.OTHER for anything still unresolved.

GPT batch completion is NOT implemented here; that step enriches the cache
before calling categorize_transactions with the updated cache.

Precedence: seed > cache > Other fallback.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, replace
from decimal import Decimal
from pathlib import Path
from typing import Optional

import openpyxl

from ledgerflow.config import Config
from ledgerflow.models import Category, ClassificationSource, Transaction
from ledgerflow.utils.normalization import normalize_merchant


# ------------------------------------------------------------------ #
# Fuel convenience rule constants                                       #
# ------------------------------------------------------------------ #

# Transactions at fuel-station merchants below this amount are treated as
# convenience-store purchases and recategorised to Leisure.
_FUEL_AMOUNT_THRESHOLD = Decimal("50")

# Transit-related patterns that take priority over fuel detection.
# If any of these appear in the normalised merchant name the fuel rule is
# not applied even if a fuel keyword is also present.
_TRANSIT_EXCEPTION_PATTERNS: frozenset[str] = frozenset({
    "רכבת",       # train / Israel Railways
    "אוטובוס",    # bus
    "רב קו",      # Rav-Kav transit card
    "רב-קו",
    "כביש",       # road / toll
    "אגרה",       # fee / toll
    "פנגו",       # Pango parking
    "סלופארק",    # SloParks
    "חניה",       # parking
    "רישיון",     # licence renewal
    "רשיון",
})

# Normalised substrings that indicate a fuel-station merchant.
_FUEL_STATION_PATTERNS: frozenset[str] = frozenset({
    "פז",
    "yellow",
    "דלק",
    "סונול",
    "דור אלון",
    "אלונית",
})


def _is_fuel_station_merchant(normalized: str) -> bool:
    """
    Return True when *normalized* (output of :func:`normalize_merchant`)
    looks like a fuel-station or fuel-station convenience-store merchant.

    Transit-related patterns (train, bus, toll, parking, licence) take
    priority: if any appear, False is returned unconditionally.
    """
    for pattern in _TRANSIT_EXCEPTION_PATTERNS:
        if pattern in normalized:
            return False
    for pattern in _FUEL_STATION_PATTERNS:
        if pattern in normalized:
            return True
    return False


# ------------------------------------------------------------------ #
# Result type                                                          #
# ------------------------------------------------------------------ #


@dataclass(frozen=True)
class CategorizationResult:
    """
    Outcome of a single categorize_transactions call.

    Attributes:
        transactions:        New Transaction objects with category filled in.
        seed_count:          Number of transactions resolved by the seed mapping.
        cache_count:         Number of transactions resolved by the cache.
        fallback_count:      Number of transactions that fell back to Other.
        unresolved_merchants: Sorted list of normalized merchant names that had
                              no valid mapping — i.e. fell back to Other because
                              neither seed nor cache held an allowed category.
                              These are the candidates for the GPT step.
    """

    transactions: list[Transaction]
    seed_count: int
    cache_count: int
    fallback_count: int
    unresolved_merchants: list[str]


# ------------------------------------------------------------------ #
# Seed mapping                                                         #
# ------------------------------------------------------------------ #

# Required column headers in the seed Excel file (matched case-insensitively
# after stripping whitespace).
_SEED_HEADERS: dict[str, str] = {
    "merchantkey": "MerchantKey",
    "merchantoriginal": "MerchantOriginal",
    "category": "Category",
}


def load_seed_mapping(path: Path) -> dict[str, Category]:
    """
    Load a merchant → category mapping from a seed Excel file.

    The first sheet of the workbook is used.  The function scans row-by-row
    for the first row that contains all three required headers
    (case-insensitive): ``MerchantKey``, ``MerchantOriginal``, ``Category``.

    For each data row below the header, both the normalized ``MerchantKey``
    and the normalized ``MerchantOriginal`` are added as keys (so either form
    can match during categorization).

    Rows with unrecognized category values are skipped silently.

    Args:
        path: Path to the seed Excel file (.xlsx / .xlsm).

    Returns:
        Dict keyed by normalized merchant name → Category.

    Raises:
        ValueError: if any of the three required columns is absent.
        FileNotFoundError: propagated from openpyxl if the file does not exist.
    """
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    try:
        ws = wb.active
        rows: list[tuple] = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()

    key_col, original_col, category_col, header_idx = _find_seed_header(rows, path)

    mapping: dict[str, Category] = {}

    for row in rows[header_idx + 1 :]:
        max_needed = max(key_col, original_col, category_col)
        if len(row) <= max_needed:
            continue

        raw_category = row[category_col]
        if raw_category is None:
            continue

        category = _parse_category(str(raw_category).strip())
        if category is None:
            continue  # unrecognized value — skip row

        raw_key = row[key_col]
        raw_original = row[original_col]

        if raw_key is not None:
            norm = normalize_merchant(str(raw_key))
            if norm:
                mapping[norm] = category

        if raw_original is not None:
            norm = normalize_merchant(str(raw_original))
            if norm:
                mapping[norm] = category

    return mapping


def _find_seed_header(
    rows: list[tuple],
    path: Path,
) -> tuple[int, int, int, int]:
    """
    Scan *rows* for the first row containing all three required seed headers.

    Returns ``(key_col, original_col, category_col, header_row_index)``.

    Raises:
        ValueError: if the headers are not found.
    """
    key_col: Optional[int] = None
    original_col: Optional[int] = None
    category_col: Optional[int] = None

    for row_idx, row in enumerate(rows):
        for col_idx, cell in enumerate(row):
            if cell is None:
                continue
            canon = str(cell).strip().lower()
            if key_col is None and canon == "merchantkey":
                key_col = col_idx
            elif original_col is None and canon == "merchantoriginal":
                original_col = col_idx
            elif category_col is None and canon == "category":
                category_col = col_idx

        if key_col is not None and original_col is not None and category_col is not None:
            return key_col, original_col, category_col, row_idx

    missing = [
        label
        for col, label in [
            (key_col, "MerchantKey"),
            (original_col, "MerchantOriginal"),
            (category_col, "Category"),
        ]
        if col is None
    ]
    raise ValueError(
        f"Seed mapping file {path.name!r} is missing required column(s): "
        f"{', '.join(missing)}. "
        f"Expected headers (case-insensitive): MerchantKey, MerchantOriginal, Category."
    )


# ------------------------------------------------------------------ #
# Cache I/O                                                            #
# ------------------------------------------------------------------ #


def load_cache(path: Path) -> dict[str, Category]:
    """
    Load the category cache from a JSON file.

    The cache is keyed by normalized merchant name.  Entries whose value is
    not a recognised Category string are silently dropped — they will be
    re-classified by the GPT step or fall back to Other.

    Returns an empty dict if the file does not exist.
    """
    if not path.exists():
        return {}
    with path.open("r", encoding="utf-8") as fh:
        raw: dict[str, str] = json.load(fh)
    result: dict[str, Category] = {}
    for merchant, cat_str in raw.items():
        category = _parse_category(cat_str)
        if category is not None:
            result[merchant] = category
    return result


def save_cache(mapping: dict[str, Category], path: Path) -> None:
    """Persist *mapping* to *path* as JSON, creating parent directories as needed."""
    path.parent.mkdir(parents=True, exist_ok=True)
    serialized = {merchant: cat.value for merchant, cat in mapping.items()}
    with path.open("w", encoding="utf-8") as fh:
        json.dump(serialized, fh, ensure_ascii=False, indent=2)


# ------------------------------------------------------------------ #
# Categorization                                                       #
# ------------------------------------------------------------------ #


def categorize_transactions(
    transactions: list[Transaction],
    seed_mapping: dict[str, Category],
    cache: dict[str, Category],
    config: Config,
    gpt_merchants: Optional[set[str]] = None,
) -> CategorizationResult:
    """
    Assign a category to every transaction and return a :class:`CategorizationResult`.

    Lookup order for each transaction (first valid match wins):
    1. **Seed mapping** — highest trust; user-curated.
    2. **Cache** — GPT-produced; lower trust than seed.
    3. **Fallback** — ``Category.OTHER``.

    A category found in seed or cache is accepted only when it belongs to
    ``config.allowed_categories()``.  A Mortgage entry is therefore rejected
    when ``config.include_mortgage`` is ``False``, and the transaction falls
    through to the next layer.

    Input transactions are never mutated.  The result contains new
    :class:`~ledgerflow.models.Transaction` objects created via
    ``dataclasses.replace``.

    Args:
        transactions:  Transactions to categorize (typically after dedup + month filter).
        seed_mapping:  Loaded from :func:`load_seed_mapping`.
        cache:         Loaded from :func:`load_cache`.
        config:        Drives the allowed-category set.

    Returns:
        :class:`CategorizationResult` with categorized transactions and stats.
    """
    allowed: set[Category] = set(config.allowed_categories())

    result_txs: list[Transaction] = []
    seed_count = cache_count = fallback_count = 0
    unresolved: set[str] = set()
    gpt_keys = gpt_merchants if gpt_merchants is not None else set()

    for tx in transactions:
        key = tx.normalized_merchant
        category: Optional[Category] = None
        source: Optional[ClassificationSource] = None

        # --- 1. Seed ---
        seed_cat = seed_mapping.get(key)
        if seed_cat is not None and seed_cat in allowed:
            category = seed_cat
            source = ClassificationSource.SEED
            seed_count += 1

        # --- 2. Cache ---
        if category is None:
            cache_cat = cache.get(key)
            if cache_cat is not None and cache_cat in allowed:
                category = cache_cat
                source = (
                    ClassificationSource.GPT
                    if key in gpt_keys
                    else ClassificationSource.CACHE
                )
                cache_count += 1

        # --- 3. Fallback ---
        if category is None:
            category = Category.OTHER
            source = ClassificationSource.FALLBACK
            fallback_count += 1
            unresolved.add(key)

        # --- 4. Fuel convenience rule (post-categorisation correction) ---
        # A Transportation-categorised transaction below the threshold at a
        # fuel-station merchant is most likely a convenience-store purchase.
        # This overrides whatever assigned Transportation (seed, cache, fallback).
        if (
            category == Category.TRANSPORTATION
            and tx.amount < _FUEL_AMOUNT_THRESHOLD
            and _is_fuel_station_merchant(tx.normalized_merchant)
        ):
            category = Category.LEISURE
            source = ClassificationSource.RULE

        result_txs.append(
            replace(tx, category=category, classification_source=source)
        )

    return CategorizationResult(
        transactions=result_txs,
        seed_count=seed_count,
        cache_count=cache_count,
        fallback_count=fallback_count,
        unresolved_merchants=sorted(unresolved),
    )


# ------------------------------------------------------------------ #
# Utilities                                                            #
# ------------------------------------------------------------------ #


def merge_mappings(*mappings: dict[str, Category]) -> dict[str, Category]:
    """
    Merge multiple merchant → category dicts; later arguments override earlier ones.

    Useful for combining GPT results into the cache before persisting.
    Note: ``categorize_transactions`` applies seed and cache separately (seed
    takes precedence), so this function is not used internally there.
    """
    merged: dict[str, Category] = {}
    for m in mappings:
        merged.update(m)
    return merged


def uncategorized_merchants(
    transactions: list[Transaction],
    mapping: dict[str, Category],
) -> list[str]:
    """
    Return a sorted list of normalized merchant names not present in *mapping*.

    Used before calling the GPT step to build the batch of unknowns.
    Pass the merged seed + cache as *mapping* to get only merchants that
    neither source knows about.
    """
    missing: set[str] = set()
    for tx in transactions:
        if tx.normalized_merchant not in mapping:
            missing.add(tx.normalized_merchant)
    return sorted(missing)


# ------------------------------------------------------------------ #
# Private helpers                                                      #
# ------------------------------------------------------------------ #


def _parse_category(value: str) -> Optional[Category]:
    """Return the matching Category for *value*, or None if unrecognised."""
    try:
        return Category(value)
    except ValueError:
        return None
