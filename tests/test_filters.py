"""Tests for ledgerflow.filters."""
from __future__ import annotations

import pytest
from datetime import date
from decimal import Decimal
from pathlib import Path

from ledgerflow.filters import filter_by_month
from ledgerflow.models import Transaction


def tx(month: int, day: int = 1, merchant: str = "shop") -> Transaction:
    return Transaction(
        date=date(2024, month, day),
        merchant=merchant,
        normalized_merchant=merchant,
        amount=Decimal("10.00"),
        source_file=Path("bank.xlsx"),
        sheet_name="Sheet1",
    )


class TestFilterByMonth:
    def test_keeps_matching_month(self):
        result = filter_by_month([tx(3)], 3)
        assert len(result) == 1

    def test_excludes_other_months(self):
        txs = [tx(1), tx(2), tx(3), tx(12)]
        result = filter_by_month(txs, 3)
        assert len(result) == 1
        assert result[0].date.month == 3

    def test_returns_empty_when_no_match(self):
        txs = [tx(1), tx(2)]
        assert filter_by_month(txs, 6) == []

    def test_empty_input_returns_empty(self):
        assert filter_by_month([], 3) == []

    def test_preserves_order(self):
        txs = [tx(3, day=15), tx(3, day=1), tx(3, day=8)]
        result = filter_by_month(txs, 3)
        days = [t.date.day for t in result]
        assert days == [15, 1, 8]

    def test_all_months_valid(self):
        all_months = [tx(m) for m in range(1, 13)]
        for m in range(1, 13):
            result = filter_by_month(all_months, m)
            assert len(result) == 1
            assert result[0].date.month == m

    def test_multiple_transactions_same_month(self):
        txs = [tx(3, day=1), tx(3, day=15), tx(4, day=1)]
        result = filter_by_month(txs, 3)
        assert len(result) == 2

    def test_matches_parsed_two_digit_dot_dates(self):
        from ledgerflow.utils.dates import parse_date

        parsed = parse_date("28.03.26")
        assert parsed is not None

        txs = [
            Transaction(
                date=parsed,
                merchant="shop",
                normalized_merchant="shop",
                amount=Decimal("10.00"),
                source_file=Path("bank.xlsx"),
                sheet_name="Sheet1",
            )
        ]
        result = filter_by_month(txs, 3)
        assert len(result) == 1

    def test_does_not_mutate_input(self):
        txs = [tx(3), tx(4)]
        filter_by_month(txs, 3)
        assert len(txs) == 2

    def test_boundary_month_january(self):
        result = filter_by_month([tx(1), tx(2)], 1)
        assert len(result) == 1

    def test_boundary_month_december(self):
        result = filter_by_month([tx(11), tx(12)], 12)
        assert len(result) == 1

    def test_invalid_month_zero_raises(self):
        with pytest.raises(ValueError, match="1 and 12"):
            filter_by_month([], 0)

    def test_invalid_month_thirteen_raises(self):
        with pytest.raises(ValueError, match="1 and 12"):
            filter_by_month([], 13)

    def test_invalid_month_negative_raises(self):
        with pytest.raises(ValueError):
            filter_by_month([], -1)

    def test_invalid_month_string_raises(self):
        with pytest.raises((ValueError, TypeError)):
            filter_by_month([], "march")  # type: ignore[arg-type]
