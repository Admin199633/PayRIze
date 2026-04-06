"""
Pipeline orchestration — ties every layer together into a single run.

The core function ``run_pipeline`` accepts pre-loaded seed and cache mappings
so it can be called from tests with plain in-memory dicts.  The convenience
wrapper ``run_pipeline_from_config`` handles loading those from ``Config``
paths and is what the UI layer will call.

Full flow
---------
extracted transactions (input)
    └─► deduplicate
    └─► filter_by_month          ── if empty: return early, no GPT
    └─► categorize (seed + cache)  [first pass]
    └─► if unresolved and gpt_client provided:
            categorize_with_gpt(unresolved_merchants)
            merge GPT results into cache (in memory)
            save_cache if config.cache_path is set
            categorize again (seed + updated cache)  [second pass]
    └─► write_raw_csv       if raw_csv_path given
    └─► write_categorized_csv  if categorized_csv_path given
    └─► return PipelineResult

No yearly workbook writing.  No UI.  No tkinter.
"""
from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

from ledgerflow.categorization import (
    categorize_transactions,
    load_cache,
    load_seed_mapping,
    merge_mappings,
    save_cache,
)
from ledgerflow.config import Config
from ledgerflow.csv_writer import write_categorized_csv, write_raw_csv
from ledgerflow.dedup import deduplicate
from ledgerflow.filters import filter_by_month
from ledgerflow.gpt_categorizer import categorize_with_gpt
from ledgerflow.models import Category, Transaction

logger = logging.getLogger(__name__)


# ------------------------------------------------------------------ #
# Result type                                                          #
# ------------------------------------------------------------------ #


@dataclass(frozen=True)
class PipelineResult:
    """
    Full statistics and output for a single pipeline run.

    Attributes:
        transactions:           Final categorized transactions (after all passes).
        extracted_count:        Total transactions passed in (before dedup).
        after_dedup_count:      Transactions remaining after deduplication.
        after_month_filter_count: Transactions remaining after month filter.
        seed_count:             Transactions resolved by the seed mapping.
        cache_count:            Transactions resolved by cache (includes GPT entries).
        gpt_count:              Transactions newly resolved by GPT in the second pass.
        fallback_count:         Transactions that fell back to Other after all passes.
        unresolved_merchants:   Normalized merchant names still at Other after all passes.
        cache_was_updated:      True if the cache file was written during this run.
    """

    transactions: list[Transaction]
    extracted_count: int
    after_dedup_count: int
    after_month_filter_count: int
    seed_count: int
    cache_count: int
    gpt_count: int
    fallback_count: int
    unresolved_merchants: list[str]
    cache_was_updated: bool

    @property
    def total_resolved(self) -> int:
        """Transactions with a non-fallback category (seed + cache + GPT)."""
        return self.seed_count + self.cache_count

    @property
    def resolution_rate(self) -> float:
        """Fraction of month-filtered transactions that were resolved (0.0–1.0)."""
        if self.after_month_filter_count == 0:
            return 1.0
        return self.total_resolved / self.after_month_filter_count


# ------------------------------------------------------------------ #
# Core pipeline (fully injectable — no file I/O beyond CSV output)    #
# ------------------------------------------------------------------ #


def run_pipeline(
    transactions: list[Transaction],
    month: int,
    seed_mapping: dict[str, Category],
    cache: dict[str, Category],
    config: Config,
    gpt_client: Any | None = None,
    raw_csv_path: Optional[Path] = None,
    categorized_csv_path: Optional[Path] = None,
) -> PipelineResult:
    """
    Run the full categorization pipeline on *transactions*.

    Args:
        transactions:         Pre-extracted list from one or more source files.
        month:                Target calendar month (1–12); transactions outside
                              this month are excluded.
        seed_mapping:         Pre-loaded seed mapping (may be empty dict).
        cache:                Pre-loaded category cache (may be empty dict).
        config:               Drives dedup precision, allowed categories, and the
                              cache-path for writing updated cache.
        gpt_client:           Injected OpenAI-compatible client.  Pass ``None``
                              to skip the GPT step entirely.
        raw_csv_path:         If provided, write the month-filtered (pre-category)
                              transactions here.
        categorized_csv_path: If provided, write the final categorized transactions
                              here.

    Returns:
        :class:`PipelineResult` with full statistics and the final transaction list.
    """
    extracted_count = len(transactions)

    # ── 1. Dedup ──────────────────────────────────────────────────────
    deduped = deduplicate(transactions, config.dedup_amount_rounding)
    after_dedup_count = len(deduped)
    logger.debug("Dedup: %d → %d transactions.", extracted_count, after_dedup_count)

    # ── 2. Month filter ───────────────────────────────────────────────
    filtered = filter_by_month(deduped, month)
    after_month_filter_count = len(filtered)
    logger.debug(
        "Month filter (month=%d): %d → %d transactions.",
        month,
        after_dedup_count,
        after_month_filter_count,
    )

    # ── 3. Early exit ─────────────────────────────────────────────────
    if not filtered:
        logger.info("No transactions for month %d after dedup; skipping categorization.", month)
        return PipelineResult(
            transactions=[],
            extracted_count=extracted_count,
            after_dedup_count=after_dedup_count,
            after_month_filter_count=0,
            seed_count=0,
            cache_count=0,
            gpt_count=0,
            fallback_count=0,
            unresolved_merchants=[],
            cache_was_updated=False,
        )

    # ── 4. First categorization pass (seed + cache, no GPT) ──────────
    first_pass = categorize_transactions(filtered, seed_mapping, cache, config)
    logger.debug(
        "First pass: seed=%d, cache=%d, fallback=%d, unresolved merchants=%d.",
        first_pass.seed_count,
        first_pass.cache_count,
        first_pass.fallback_count,
        len(first_pass.unresolved_merchants),
    )

    # ── 5. GPT enrichment ────────────────────────────────────────────
    gpt_count = 0
    cache_was_updated = False
    final_result = first_pass

    if gpt_client is not None and first_pass.unresolved_merchants:
        logger.info(
            "Calling GPT for %d unresolved merchants.",
            len(first_pass.unresolved_merchants),
        )
        gpt_cats = categorize_with_gpt(
            first_pass.unresolved_merchants, config, gpt_client
        )

        if gpt_cats:
            updated_cache = merge_mappings(cache, gpt_cats)

            # Persist cache only if a path is configured.
            if config.cache_path is not None:
                save_cache(updated_cache, config.cache_path)
                cache_was_updated = True
                logger.info("Cache updated with %d new entries.", len(gpt_cats))

            # ── 6. Second categorization pass with enriched cache ──
            second_pass = categorize_transactions(
                filtered,
                seed_mapping,
                updated_cache,
                config,
                gpt_merchants=set(gpt_cats),
            )
            gpt_count = second_pass.cache_count - first_pass.cache_count
            final_result = second_pass
            logger.debug(
                "Second pass: seed=%d, cache=%d (gpt_resolved=%d), fallback=%d.",
                second_pass.seed_count,
                second_pass.cache_count,
                gpt_count,
                second_pass.fallback_count,
            )
        else:
            logger.info("GPT returned no usable results; first-pass result kept.")

    # ── 7. CSV outputs ────────────────────────────────────────────────
    if raw_csv_path is not None:
        write_raw_csv(filtered, raw_csv_path)
        logger.info("Raw CSV written to %s (%d rows).", raw_csv_path, len(filtered))

    if categorized_csv_path is not None:
        write_categorized_csv(final_result.transactions, categorized_csv_path)
        logger.info(
            "Categorized CSV written to %s (%d rows).",
            categorized_csv_path,
            len(final_result.transactions),
        )

    return PipelineResult(
        transactions=final_result.transactions,
        extracted_count=extracted_count,
        after_dedup_count=after_dedup_count,
        after_month_filter_count=after_month_filter_count,
        seed_count=final_result.seed_count,
        cache_count=final_result.cache_count,
        gpt_count=gpt_count,
        fallback_count=final_result.fallback_count,
        unresolved_merchants=final_result.unresolved_merchants,
        cache_was_updated=cache_was_updated,
    )


# ------------------------------------------------------------------ #
# Convenience wrapper — loads seed/cache from config paths            #
# ------------------------------------------------------------------ #


def run_pipeline_from_config(
    transactions: list[Transaction],
    month: int,
    config: Config,
    gpt_client: Any | None = None,
    raw_csv_path: Optional[Path] = None,
    categorized_csv_path: Optional[Path] = None,
) -> PipelineResult:
    """
    Convenience wrapper that loads the seed mapping and cache from the paths
    stored in *config* before delegating to :func:`run_pipeline`.

    This is the entry point for the UI layer; business-logic tests should call
    ``run_pipeline`` directly with pre-loaded dicts.

    If ``config.seed_mapping_path`` is ``None``, an empty seed is used.
    If ``config.cache_path`` is ``None``, an empty cache is used (and the
    cache is never written).
    """
    seed: dict[str, Category] = (
        load_seed_mapping(config.seed_mapping_path)
        if config.seed_mapping_path is not None
        else {}
    )
    cache: dict[str, Category] = (
        load_cache(config.cache_path)
        if config.cache_path is not None
        else {}
    )
    return run_pipeline(
        transactions,
        month,
        seed,
        cache,
        config,
        gpt_client,
        raw_csv_path,
        categorized_csv_path,
    )
