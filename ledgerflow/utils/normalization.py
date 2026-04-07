"""
Merchant name normalization.

The normalized form is used as the key for deduplication and category
mapping, so it must be stable and deterministic across source files.
"""
from __future__ import annotations

import re
import unicodedata


def normalize_merchant(name: str) -> str:
    """
    Return a canonical form of a merchant name suitable for matching.

    Steps applied in order:
    1. Unicode NFKC — collapses ligatures, fullwidth chars, etc.
    2. Lowercase + strip outer whitespace.
    3. Collapse runs of internal whitespace to a single space.
    4. Remove characters that are not alphanumeric, space, hyphen, or apostrophe.
    5. Collapse whitespace again after removal.
    """
    if not name:
        return ""

    result = unicodedata.normalize("NFKC", name)
    result = result.lower().strip()
    result = re.sub(r"\s+", " ", result)
    # Keep word chars (\w includes unicode letters/digits/_), spaces, hyphens, apostrophes.
    result = re.sub(r"[^\w\s\-']", "", result)
    result = re.sub(r"\s+", " ", result).strip()
    return result


def canonicalize_header(raw: str) -> str:
    """
    Return a lowercase-stripped version of a header cell value for alias matching.
    """
    return raw.strip().lower() if raw else ""
