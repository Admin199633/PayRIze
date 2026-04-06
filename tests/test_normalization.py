import pytest
from ledgerflow.utils.normalization import normalize_merchant, canonicalize_header


class TestNormalizeMerchant:
    def test_empty_string(self):
        assert normalize_merchant("") == ""

    def test_lowercase(self):
        assert normalize_merchant("AMAZON") == "amazon"

    def test_strip_outer_whitespace(self):
        assert normalize_merchant("  Amazon  ") == "amazon"

    def test_collapse_internal_whitespace(self):
        assert normalize_merchant("Amazon   Prime") == "amazon prime"

    def test_remove_punctuation(self):
        # dots and commas stripped; letters/spaces kept
        assert normalize_merchant("Amazon.com, Inc.") == "amazoncom inc"

    def test_preserve_hyphen(self):
        assert normalize_merchant("Wal-Mart") == "wal-mart"

    def test_preserve_apostrophe(self):
        assert normalize_merchant("McDonald's") == "mcdonald's"

    def test_unicode_nfkc_fullwidth(self):
        # Fullwidth Latin letters → ASCII equivalent
        assert normalize_merchant("ＡＭＡＺＯＮ") == "amazon"

    def test_unicode_nfkc_ligature(self):
        # ﬁ ligature → fi
        assert normalize_merchant("ﬁle") == "file"

    def test_hebrew_merchant(self):
        # Hebrew text should be preserved (normalized, not stripped)
        result = normalize_merchant("  סופר פארם  ")
        assert result == "סופר פארם"

    def test_mixed_symbols(self):
        result = normalize_merchant("7-Eleven #42 (Main St.)")
        assert result == "7-eleven 42 main st"

    def test_underscore_kept(self):
        # \w includes underscore
        assert normalize_merchant("my_store") == "my_store"


class TestCanonicalizeHeader:
    def test_strips_and_lowercases(self):
        assert canonicalize_header("  Purchase Date  ") == "purchase date"

    def test_empty(self):
        assert canonicalize_header("") == ""
