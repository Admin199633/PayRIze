import pytest
from decimal import Decimal
from ledgerflow.utils.amounts import parse_amount, round_for_dedup


class TestParseAmount:
    def test_int(self):
        assert parse_amount(100) == Decimal("100")

    def test_float(self):
        assert parse_amount(19.99) == Decimal("19.99")

    def test_decimal_passthrough(self):
        d = Decimal("42.10")
        assert parse_amount(d) is d

    def test_plain_string(self):
        assert parse_amount("123.45") == Decimal("123.45")

    def test_currency_symbol(self):
        assert parse_amount("$123.45") == Decimal("123.45")

    def test_shekel_symbol(self):
        assert parse_amount("₪99.00") == Decimal("99.00")

    def test_thousands_comma(self):
        assert parse_amount("1,234.56") == Decimal("1234.56")

    def test_negative_sign(self):
        assert parse_amount("-50.00") == Decimal("-50.00")

    def test_negative_parentheses(self):
        assert parse_amount("(50.00)") == Decimal("-50.00")

    def test_negative_parentheses_with_currency(self):
        assert parse_amount("($1,200.00)") == Decimal("-1200.00")

    def test_none(self):
        assert parse_amount(None) is None

    def test_empty_string(self):
        assert parse_amount("") is None

    def test_whitespace_only(self):
        assert parse_amount("   ") is None

    def test_non_numeric_string(self):
        assert parse_amount("N/A") is None

    def test_zero(self):
        assert parse_amount("0") == Decimal("0")

    def test_zero_float(self):
        assert parse_amount(0.0) == Decimal("0.0")


class TestRoundForDedup:
    def test_round_zero_places_rounds_up(self):
        assert round_for_dedup(Decimal("123.67")) == Decimal("124")

    def test_round_zero_places_rounds_down(self):
        assert round_for_dedup(Decimal("123.40")) == Decimal("123")

    def test_round_two_places(self):
        assert round_for_dedup(Decimal("123.456"), 2) == Decimal("123.46")

    def test_round_half_up(self):
        # 0.5 should round up (ROUND_HALF_UP)
        assert round_for_dedup(Decimal("0.5")) == Decimal("1")

    def test_exact_value_unchanged(self):
        assert round_for_dedup(Decimal("100.00")) == Decimal("100")

    def test_negative_amount(self):
        assert round_for_dedup(Decimal("-123.67")) == Decimal("-124")
