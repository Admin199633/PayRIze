import pytest
from datetime import date, datetime
from ledgerflow.utils.dates import parse_date, is_valid_date, month_of


class TestParseDate:
    def test_date_object_passthrough(self):
        d = date(2024, 3, 15)
        assert parse_date(d) == d

    def test_datetime_stripped_to_date(self):
        dt = datetime(2024, 3, 15, 10, 30, 0)
        assert parse_date(dt) == date(2024, 3, 15)

    def test_iso_string(self):
        assert parse_date("2024-03-15") == date(2024, 3, 15)

    def test_slash_dd_mm_yyyy(self):
        assert parse_date("15/03/2024") == date(2024, 3, 15)

    def test_slash_mm_dd_yyyy(self):
        # Falls back to US format if dd/mm fails (month > 12 not possible, so
        # this format is tried second)
        assert parse_date("03/15/2024") == date(2024, 3, 15)

    def test_dot_format(self):
        assert parse_date("15.03.2024") == date(2024, 3, 15)

    def test_dot_format_two_digit_year(self):
        assert parse_date("28.03.26") == date(2026, 3, 28)

    def test_hyphen_format(self):
        assert parse_date("15-03-2024") == date(2024, 3, 15)

    def test_compact_format(self):
        assert parse_date("20240315") == date(2024, 3, 15)

    def test_two_digit_year(self):
        assert parse_date("15/03/24") == date(2024, 3, 15)

    def test_none_input(self):
        assert parse_date(None) is None

    def test_invalid_string(self):
        assert parse_date("not-a-date") is None

    def test_empty_string(self):
        assert parse_date("") is None

    def test_leading_trailing_spaces(self):
        assert parse_date("  2024-03-15  ") == date(2024, 3, 15)


class TestIsValidDate:
    def test_valid(self):
        assert is_valid_date("2024-01-01") is True

    def test_invalid(self):
        assert is_valid_date("garbage") is False

    def test_none(self):
        assert is_valid_date(None) is False


class TestMonthOf:
    def test_january(self):
        assert month_of(date(2024, 1, 31)) == 1

    def test_december(self):
        assert month_of(date(2024, 12, 1)) == 12
