"""
GPT-based batch categorization for merchants not resolved by seed or cache.

Architecture
------------
- No side effects: this module never reads or writes the cache file.
  The caller is responsible for merging the result into the cache and saving.
- The OpenAI client is injected so tests can pass a mock without patching.
- ``openai`` is an optional dependency (``pip install ledgerflow[gpt]``).
  It is imported lazily so the rest of the package works without it.

Typical caller flow
-------------------
  result = categorize_transactions(txs, seed, cache, config)  # first pass

  if result.unresolved_merchants:
      import openai
      client = openai.OpenAI(api_key="...")
      gpt_cats = categorize_with_gpt(result.unresolved_merchants, config, client)
      updated_cache = merge_mappings(cache, gpt_cats)
      save_cache(updated_cache, config.cache_path)
      result = categorize_transactions(txs, seed, updated_cache, config)  # second pass
"""
from __future__ import annotations

import json
import logging
from typing import Any, Optional

from ledgerflow.config import Config
from ledgerflow.models import Category

logger = logging.getLogger(__name__)

# ------------------------------------------------------------------ #
# Constants                                                            #
# ------------------------------------------------------------------ #

DEFAULT_MODEL = "gpt-4o-mini"
DEFAULT_BATCH_SIZE = 50

_SYSTEM_PROMPT_TEMPLATE = """\
You are a financial transaction categorizer.

Classify each merchant name into exactly one of these categories:
{categories}

Rules:
- Return ONLY a valid JSON object.
- Each key is a merchant name exactly as given in the input.
- Each value is one of the category names listed above (exact spelling and casing).
- If you are unsure, use "Other".
- Do not add any explanation, markdown, or extra keys.
- Merchant names may be in Hebrew or English."""


# ------------------------------------------------------------------ #
# Public API                                                           #
# ------------------------------------------------------------------ #


def categorize_with_gpt(
    merchants: list[str],
    config: Config,
    client: Any,
    model: str = DEFAULT_MODEL,
    batch_size: int = DEFAULT_BATCH_SIZE,
) -> dict[str, Category]:
    """
    Classify *merchants* via GPT and return a merchant → Category mapping.

    Merchants for which GPT returns an unrecognised category are silently
    omitted from the result — they will fall back to ``Category.OTHER`` when
    the caller runs :func:`~ledgerflow.categorization.categorize_transactions`.

    A batch that raises an API exception is logged as a warning and skipped;
    merchants from that batch also fall back to Other.

    Args:
        merchants:  Normalised merchant names (typically
                    ``CategorizationResult.unresolved_merchants``).
        config:     Provides the allowed category list for the prompt.
        client:     An ``openai.OpenAI``-compatible client (duck-typed so tests
                    can inject a mock without patching).
        model:      OpenAI model identifier.
        batch_size: Maximum number of merchants per API call.

    Returns:
        Dict of normalised merchant name → :class:`~ledgerflow.models.Category`.
    """
    if not merchants:
        return {}

    system_prompt = _build_system_prompt(config)
    result: dict[str, Category] = {}

    for batch in _make_batches(merchants, batch_size):
        try:
            batch_result = _classify_batch(batch, system_prompt, model, client)
            result.update(batch_result)
        except Exception as exc:  # noqa: BLE001
            logger.warning(
                "GPT batch failed (%d merchants): %s — those merchants will "
                "fall back to Other.",
                len(batch),
                exc,
            )

    return result


# ------------------------------------------------------------------ #
# Internal helpers — exposed for unit testing                          #
# ------------------------------------------------------------------ #


def _build_system_prompt(config: Config) -> str:
    """Build the system prompt with the allowed categories for *config*."""
    category_lines = "\n".join(
        f"- {cat.value}" for cat in config.allowed_categories()
    )
    return _SYSTEM_PROMPT_TEMPLATE.format(categories=category_lines)


def _classify_batch(
    merchants: list[str],
    system_prompt: str,
    model: str,
    client: Any,
) -> dict[str, Category]:
    """
    Send one batch of merchant names to the OpenAI chat completions endpoint
    and return a parsed merchant → Category mapping.

    Raises whatever the API client raises on network or authentication errors
    so ``categorize_with_gpt`` can log and skip the batch.
    """
    user_message = json.dumps(merchants, ensure_ascii=False)

    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message},
        ],
        temperature=0,
        response_format={"type": "json_object"},
    )

    raw_text: str = response.choices[0].message.content
    return _parse_gpt_response(raw_text)


def _parse_gpt_response(raw_text: str) -> dict[str, Category]:
    """
    Parse a GPT JSON response string into a merchant → Category dict.

    Entries with unrecognised category strings are silently dropped.

    Raises:
        ValueError: if *raw_text* is not a JSON object.
    """
    try:
        data = json.loads(raw_text)
    except json.JSONDecodeError as exc:
        raise ValueError(f"GPT response is not valid JSON: {exc}") from exc

    if not isinstance(data, dict):
        raise ValueError(
            f"GPT response must be a JSON object, got {type(data).__name__}"
        )

    result: dict[str, Category] = {}
    for merchant, cat_str in data.items():
        if not isinstance(merchant, str) or not isinstance(cat_str, str):
            logger.debug(
                "Skipping non-string entry in GPT response: %r → %r",
                merchant,
                cat_str,
            )
            continue
        category = _safe_parse_category(cat_str)
        if category is not None:
            result[merchant] = category
        else:
            logger.debug(
                "GPT returned unrecognised category %r for merchant %r — skipped.",
                cat_str,
                merchant,
            )

    return result


def _make_batches(items: list[str], batch_size: int) -> list[list[str]]:
    """Split *items* into sub-lists of at most *batch_size* elements."""
    if batch_size < 1:
        raise ValueError(f"batch_size must be >= 1, got {batch_size}")
    return [items[i : i + batch_size] for i in range(0, len(items), batch_size)]


# ------------------------------------------------------------------ #
# Private                                                              #
# ------------------------------------------------------------------ #


def _safe_parse_category(value: str) -> Optional[Category]:
    """Return the matching :class:`Category` for *value*, or ``None``."""
    try:
        return Category(value)
    except ValueError:
        return None
