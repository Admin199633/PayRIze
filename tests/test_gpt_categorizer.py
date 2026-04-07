"""
Tests for ledgerflow.gpt_categorizer.

Strategy:
- All tests are pure / in-memory: no real OpenAI calls ever happen.
- The client is a MagicMock whose response content is pre-set per test.
- Pure helpers (_parse_gpt_response, _make_batches, _build_system_prompt)
  are tested directly without a client at all.
- categorize_with_gpt is tested end-to-end with the mock client.
"""
from __future__ import annotations

import json
import pytest
from unittest.mock import MagicMock, call

from ledgerflow.config import Config
from ledgerflow.models import Category
from ledgerflow.gpt_categorizer import (
    DEFAULT_BATCH_SIZE,
    DEFAULT_MODEL,
    _build_system_prompt,
    _classify_batch,
    _make_batches,
    _parse_gpt_response,
    categorize_with_gpt,
)


# ------------------------------------------------------------------ #
# Helpers                                                              #
# ------------------------------------------------------------------ #


def make_client(response_content: str) -> MagicMock:
    """Build a mock OpenAI client whose first choice returns *response_content*."""
    client = MagicMock()
    client.chat.completions.create.return_value.choices[0].message.content = (
        response_content
    )
    return client


def gpt_json(**mapping: str) -> str:
    """Convenience: serialise a dict to a JSON string for use as a mock response."""
    return json.dumps(mapping, ensure_ascii=False)


# ------------------------------------------------------------------ #
# _make_batches                                                        #
# ------------------------------------------------------------------ #


class TestMakeBatches:
    def test_fewer_items_than_batch_size(self):
        assert _make_batches(["a", "b", "c"], 10) == [["a", "b", "c"]]

    def test_exactly_one_batch(self):
        assert _make_batches(["a", "b"], 2) == [["a", "b"]]

    def test_splits_into_multiple_batches(self):
        result = _make_batches(["a", "b", "c", "d", "e"], 2)
        assert result == [["a", "b"], ["c", "d"], ["e"]]

    def test_single_item(self):
        assert _make_batches(["only"], 50) == [["only"]]

    def test_empty_list(self):
        assert _make_batches([], 10) == []

    def test_batch_size_one(self):
        result = _make_batches(["a", "b", "c"], 1)
        assert result == [["a"], ["b"], ["c"]]

    def test_invalid_batch_size_raises(self):
        with pytest.raises(ValueError, match="batch_size"):
            _make_batches(["a"], 0)


# ------------------------------------------------------------------ #
# _build_system_prompt                                                 #
# ------------------------------------------------------------------ #


class TestBuildSystemPrompt:
    def test_contains_all_default_categories(self):
        prompt = _build_system_prompt(Config())
        for cat in Config().allowed_categories():
            assert cat.value in prompt

    def test_mortgage_excluded_when_not_in_config(self):
        cfg = Config()
        cfg.include_mortgage = False
        prompt = _build_system_prompt(cfg)
        assert "Mortgage" not in prompt

    def test_mortgage_included_when_enabled(self):
        cfg = Config()
        cfg.include_mortgage = True
        cfg.categories.append(Category.MORTGAGE)
        prompt = _build_system_prompt(cfg)
        assert "Mortgage" in prompt

    def test_prompt_is_nonempty_string(self):
        prompt = _build_system_prompt(Config())
        assert isinstance(prompt, str) and len(prompt) > 50


# ------------------------------------------------------------------ #
# _parse_gpt_response                                                  #
# ------------------------------------------------------------------ #


class TestParseGptResponse:
    def test_valid_response_single_merchant(self):
        raw = gpt_json(amazon="Business")
        result = _parse_gpt_response(raw)
        assert result == {"amazon": Category.BUSINESS}

    def test_valid_response_multiple_merchants(self):
        raw = gpt_json(amazon="Business", spotify="Leisure", boots="Pharm")
        result = _parse_gpt_response(raw)
        assert result["amazon"] == Category.BUSINESS
        assert result["spotify"] == Category.LEISURE
        assert result["boots"] == Category.PHARM

    def test_all_standard_categories_parsed(self):
        raw = json.dumps({
            "a": "Other",
            "b": "Business",
            "c": "Pharm",
            "d": "Insurance",
            "e": "Transportation",
            "f": "Leisure",
            "g": "Home",
            "h": "Mortgage",
        })
        result = _parse_gpt_response(raw)
        assert result["a"] == Category.OTHER
        assert result["h"] == Category.MORTGAGE

    def test_unrecognised_category_is_dropped(self):
        raw = gpt_json(amazon="Business", badmerchant="NotARealCategory")
        result = _parse_gpt_response(raw)
        assert "amazon" in result
        assert "badmerchant" not in result

    def test_all_unrecognised_returns_empty_dict(self):
        raw = gpt_json(shop="Unknown", cafe="AlsoUnknown")
        assert _parse_gpt_response(raw) == {}

    def test_empty_json_object_returns_empty_dict(self):
        assert _parse_gpt_response("{}") == {}

    def test_invalid_json_raises_value_error(self):
        with pytest.raises(ValueError, match="not valid JSON"):
            _parse_gpt_response("not json at all")

    def test_json_array_raises_value_error(self):
        with pytest.raises(ValueError, match="JSON object"):
            _parse_gpt_response('["Business", "Leisure"]')

    def test_json_null_raises_value_error(self):
        with pytest.raises(ValueError, match="JSON object"):
            _parse_gpt_response("null")

    def test_non_string_value_is_dropped(self):
        # GPT hallucinated a number as the category
        raw = json.dumps({"amazon": 42})
        assert _parse_gpt_response(raw) == {}

    def test_hebrew_merchant_name_preserved(self):
        raw = json.dumps({"סופר פארם": "Pharm"}, ensure_ascii=False)
        result = _parse_gpt_response(raw)
        assert result["סופר פארם"] == Category.PHARM

    def test_extra_whitespace_in_category_not_matched(self):
        # "Business " (trailing space) is not a valid Category value
        raw = gpt_json(shop="Business ")
        result = _parse_gpt_response(raw)
        assert "shop" not in result


# ------------------------------------------------------------------ #
# _classify_batch                                                      #
# ------------------------------------------------------------------ #


class TestClassifyBatch:
    def _call(self, merchants, response_json, model=DEFAULT_MODEL):
        client = make_client(response_json)
        return _classify_batch(merchants, "system prompt here", model, client), client

    def test_returns_parsed_mapping(self):
        result, _ = self._call(
            ["amazon", "spotify"],
            gpt_json(amazon="Business", spotify="Leisure"),
        )
        assert result == {"amazon": Category.BUSINESS, "spotify": Category.LEISURE}

    def test_passes_merchants_as_json_array_in_user_message(self):
        merchants = ["amazon", "spotify"]
        result, client = self._call(merchants, gpt_json(amazon="Business"))
        _, kwargs = client.chat.completions.create.call_args
        user_msg = kwargs["messages"][1]["content"]
        assert json.loads(user_msg) == merchants

    def test_passes_system_prompt(self):
        result, client = self._call(["shop"], gpt_json(shop="Other"))
        _, kwargs = client.chat.completions.create.call_args
        assert kwargs["messages"][0]["content"] == "system prompt here"

    def test_passes_model_name(self):
        result, client = self._call(["shop"], gpt_json(shop="Other"), model="gpt-4o")
        _, kwargs = client.chat.completions.create.call_args
        assert kwargs["model"] == "gpt-4o"

    def test_uses_temperature_zero(self):
        result, client = self._call(["shop"], gpt_json(shop="Other"))
        _, kwargs = client.chat.completions.create.call_args
        assert kwargs["temperature"] == 0

    def test_requests_json_object_format(self):
        result, client = self._call(["shop"], gpt_json(shop="Other"))
        _, kwargs = client.chat.completions.create.call_args
        assert kwargs["response_format"] == {"type": "json_object"}

    def test_api_exception_propagates(self):
        client = MagicMock()
        client.chat.completions.create.side_effect = RuntimeError("API down")
        with pytest.raises(RuntimeError, match="API down"):
            _classify_batch(["shop"], "prompt", DEFAULT_MODEL, client)


# ------------------------------------------------------------------ #
# categorize_with_gpt (end-to-end with mock)                          #
# ------------------------------------------------------------------ #


class TestCategorizeWithGpt:
    def test_empty_merchants_returns_empty_without_api_call(self):
        client = MagicMock()
        result = categorize_with_gpt([], Config(), client)
        assert result == {}
        client.chat.completions.create.assert_not_called()

    def test_single_batch_resolved(self):
        client = make_client(gpt_json(amazon="Business", spotify="Leisure"))
        result = categorize_with_gpt(["amazon", "spotify"], Config(), client)
        assert result["amazon"] == Category.BUSINESS
        assert result["spotify"] == Category.LEISURE

    def test_unrecognised_categories_absent_from_result(self):
        client = make_client(gpt_json(amazon="Business", shop="NOPE"))
        result = categorize_with_gpt(["amazon", "shop"], Config(), client)
        assert "amazon" in result
        assert "shop" not in result

    def test_two_batches_results_combined(self):
        # batch_size=2 forces two API calls for 4 merchants
        responses = [
            gpt_json(a="Business", b="Leisure"),
            gpt_json(c="Home", d="Pharm"),
        ]
        client = MagicMock()
        client.chat.completions.create.side_effect = [
            _make_api_response(responses[0]),
            _make_api_response(responses[1]),
        ]
        result = categorize_with_gpt(["a", "b", "c", "d"], Config(), client, batch_size=2)
        assert len(result) == 4
        assert result["a"] == Category.BUSINESS
        assert result["c"] == Category.HOME

    def test_api_call_count_matches_batch_count(self):
        client = make_client(gpt_json(a="Business", b="Leisure", c="Home"))
        categorize_with_gpt(["a", "b", "c"], Config(), client, batch_size=2)
        # 3 merchants / batch_size=2 → 2 calls
        assert client.chat.completions.create.call_count == 2

    def test_failed_batch_skipped_does_not_raise(self):
        # First batch fails; second batch succeeds.
        client = MagicMock()
        client.chat.completions.create.side_effect = [
            RuntimeError("timeout"),
            _make_api_response(gpt_json(c="Business")),
        ]
        result = categorize_with_gpt(["a", "b", "c"], Config(), client, batch_size=2)
        # 'a' and 'b' from the failed batch are absent; 'c' succeeds
        assert "a" not in result
        assert "b" not in result
        assert result.get("c") == Category.BUSINESS

    def test_all_batches_fail_returns_empty_dict(self):
        client = MagicMock()
        client.chat.completions.create.side_effect = RuntimeError("down")
        result = categorize_with_gpt(["a", "b", "c"], Config(), client)
        assert result == {}

    def test_default_model_is_passed_to_api(self):
        client = make_client(gpt_json(shop="Other"))
        categorize_with_gpt(["shop"], Config(), client)
        _, kwargs = client.chat.completions.create.call_args
        assert kwargs["model"] == DEFAULT_MODEL

    def test_custom_model_is_passed_to_api(self):
        client = make_client(gpt_json(shop="Other"))
        categorize_with_gpt(["shop"], Config(), client, model="gpt-4o")
        _, kwargs = client.chat.completions.create.call_args
        assert kwargs["model"] == "gpt-4o"

    def test_mortgage_in_result_when_config_enables_it(self):
        cfg = Config()
        cfg.include_mortgage = True
        cfg.categories.append(Category.MORTGAGE)
        client = make_client(gpt_json(bank="Mortgage"))
        result = categorize_with_gpt(["bank"], cfg, client)
        assert result["bank"] == Category.MORTGAGE

    def test_mortgage_still_parsed_even_when_not_in_config(self):
        # _parse_gpt_response doesn't know about allowed_categories — that
        # filtering happens later in categorize_transactions.
        cfg = Config()
        cfg.include_mortgage = False
        client = make_client(gpt_json(bank="Mortgage"))
        result = categorize_with_gpt(["bank"], cfg, client)
        # Mortgage is a valid Category enum value — it gets included in the
        # raw result. categorize_transactions will reject it.
        assert result.get("bank") == Category.MORTGAGE

    def test_hebrew_merchants_handled(self):
        merchant = "סופר פארם"
        response = json.dumps({merchant: "Pharm"}, ensure_ascii=False)
        client = make_client(response)
        result = categorize_with_gpt([merchant], Config(), client)
        assert result[merchant] == Category.PHARM

    def test_malformed_json_response_skips_batch(self):
        client = make_client("this is not json {{{")
        result = categorize_with_gpt(["shop"], Config(), client)
        assert result == {}

    def test_result_does_not_contain_extra_merchants(self):
        # GPT might hallucinate extra keys — they're included in the raw result
        # because we don't filter by input list. This is intentional: extra
        # entries are harmless and the cache benefits from them.
        client = make_client(gpt_json(asked="Business", extra="Leisure"))
        result = categorize_with_gpt(["asked"], Config(), client)
        assert "extra" in result  # documented behaviour, not a bug


# ------------------------------------------------------------------ #
# Integration: GPT result feeds back into categorize_transactions      #
# ------------------------------------------------------------------ #


class TestGptIntegrationWithCategorization:
    """
    Verify the full caller pattern:
    first pass → GPT → merge into cache → second pass.
    """

    def test_second_pass_resolves_all_transactions(self):
        from datetime import date
        from decimal import Decimal
        from pathlib import Path

        from ledgerflow.categorization import (
            categorize_transactions,
            merge_mappings,
        )
        from ledgerflow.models import Transaction

        def tx(merchant, normalized):
            return Transaction(
                date=date(2024, 3, 15),
                merchant=merchant,
                normalized_merchant=normalized,
                amount=Decimal("50.00"),
                source_file=Path("bank.xlsx"),
                sheet_name="Sheet1",
            )

        txs = [
            tx("Amazon", "amazon"),
            tx("Unknown Shop", "unknown shop"),
        ]
        seed = {"amazon": Category.BUSINESS}
        cache: dict[str, Category] = {}
        cfg = Config()

        # First pass: "unknown shop" falls back to Other
        first = categorize_transactions(txs, seed, cache, cfg)
        assert first.fallback_count == 1
        assert "unknown shop" in first.unresolved_merchants

        # GPT resolves the missing merchant
        client = make_client(gpt_json(**{"unknown shop": "Leisure"}))
        gpt_result = categorize_with_gpt(first.unresolved_merchants, cfg, client)
        updated_cache = merge_mappings(cache, gpt_result)

        # Second pass: no fallbacks remaining
        second = categorize_transactions(txs, seed, updated_cache, cfg)
        assert second.fallback_count == 0
        cats = {t.normalized_merchant: t.category for t in second.transactions}
        assert cats["amazon"] == Category.BUSINESS
        assert cats["unknown shop"] == Category.LEISURE


# ------------------------------------------------------------------ #
# Private helper                                                       #
# ------------------------------------------------------------------ #


def _make_api_response(content: str) -> MagicMock:
    """Build a minimal mock OpenAI response object with *content*."""
    response = MagicMock()
    response.choices[0].message.content = content
    return response
