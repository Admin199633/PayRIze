import pytest
from datetime import date
from decimal import Decimal
from pathlib import Path

from ledgerflow.models import Transaction, Category
from ledgerflow.dedup import make_dedup_key, deduplicate, find_duplicates


# ------------------------------------------------------------------ #
# Fixtures / helpers                                                   #
# ------------------------------------------------------------------ #

def tx(
    d: str = "2024-03-15",
    merchant: str = "amazon",
    normalized: str | None = None,
    amount: str = "100.00",
    source: str = "file1.xlsx",
    sheet: str = "Sheet1",
) -> Transaction:
    return Transaction(
        date=date.fromisoformat(d),
        merchant=merchant,
        normalized_merchant=normalized if normalized is not None else merchant.lower().strip(),
        amount=Decimal(amount),
        source_file=Path(source),
        sheet_name=sheet,
    )


# ------------------------------------------------------------------ #
# make_dedup_key                                                        #
# ------------------------------------------------------------------ #

class TestMakeDedupKey:
    def test_basic_key(self):
        t = tx()
        key = make_dedup_key(t)
        assert key == (date(2024, 3, 15), "amazon", Decimal("100"))

    def test_key_rounds_amount(self):
        t = tx(amount="100.49")
        key = make_dedup_key(t, amount_decimal_places=0)
        assert key[2] == Decimal("100")

    def test_key_respects_decimal_places(self):
        t = tx(amount="100.456")
        key = make_dedup_key(t, amount_decimal_places=2)
        assert key[2] == Decimal("100.46")

    def test_key_uses_normalized_merchant(self):
        t = tx(merchant="AMAZON.COM", normalized="amazon com")
        key = make_dedup_key(t)
        assert key[1] == "amazon com"


# ------------------------------------------------------------------ #
# deduplicate                                                           #
# ------------------------------------------------------------------ #

class TestDeduplicate:
    def test_empty_list(self):
        assert deduplicate([]) == []

    def test_no_duplicates(self):
        transactions = [tx(merchant="amazon"), tx(merchant="walmart")]
        result = deduplicate(transactions)
        assert len(result) == 2

    def test_exact_duplicate_removed(self):
        t1 = tx()
        t2 = tx()
        result = deduplicate([t1, t2])
        assert len(result) == 1
        assert result[0] is t1  # first occurrence wins

    def test_duplicate_from_different_source_file(self):
        t1 = tx(source="bank_a.xlsx")
        t2 = tx(source="bank_b.xlsx")
        result = deduplicate([t1, t2])
        assert len(result) == 1

    def test_different_dates_kept(self):
        t1 = tx(d="2024-03-15")
        t2 = tx(d="2024-03-16")
        assert len(deduplicate([t1, t2])) == 2

    def test_different_amounts_kept(self):
        t1 = tx(amount="100.00")
        t2 = tx(amount="200.00")
        assert len(deduplicate([t1, t2])) == 2

    def test_different_merchants_kept(self):
        t1 = tx(merchant="amazon")
        t2 = tx(merchant="walmart")
        assert len(deduplicate([t1, t2])) == 2

    def test_preserves_relative_order(self):
        t1 = tx(merchant="amazon")
        t2 = tx(merchant="walmart")
        t3 = tx(merchant="amazon")  # duplicate of t1
        result = deduplicate([t1, t2, t3])
        assert [r.merchant for r in result] == ["amazon", "walmart"]

    def test_rounding_merges_close_amounts(self):
        # 100.20 and 100.49 both round to 100 at 0 decimal places → duplicate
        t1 = tx(amount="100.20")
        t2 = tx(amount="100.49")
        result = deduplicate([t1, t2], amount_decimal_places=0)
        assert len(result) == 1

    def test_rounding_keeps_different_amounts_at_higher_precision(self):
        t1 = tx(amount="100.30")
        t2 = tx(amount="100.70")
        result = deduplicate([t1, t2], amount_decimal_places=1)
        assert len(result) == 2

    def test_category_is_preserved(self):
        t1 = tx()
        t1.category = Category.BUSINESS
        t2 = tx()
        result = deduplicate([t1, t2])
        assert result[0].category == Category.BUSINESS


# ------------------------------------------------------------------ #
# find_duplicates                                                       #
# ------------------------------------------------------------------ #

class TestFindDuplicates:
    def test_no_duplicates_returns_empty(self):
        transactions = [tx(merchant="amazon"), tx(merchant="walmart")]
        assert find_duplicates(transactions) == {}

    def test_finds_one_duplicate_group(self):
        t1 = tx()
        t2 = tx()
        t3 = tx(merchant="walmart")
        groups = find_duplicates([t1, t2, t3])
        assert len(groups) == 1
        group = next(iter(groups.values()))
        assert len(group) == 2

    def test_finds_multiple_groups(self):
        t1 = tx(merchant="amazon")
        t2 = tx(merchant="amazon")
        t3 = tx(merchant="walmart")
        t4 = tx(merchant="walmart")
        groups = find_duplicates([t1, t2, t3, t4])
        assert len(groups) == 2

    def test_key_structure(self):
        t1 = tx()
        t2 = tx()
        groups = find_duplicates([t1, t2])
        key = next(iter(groups.keys()))
        # key is (date, str, Decimal)
        assert isinstance(key[0], date)
        assert isinstance(key[1], str)
        assert isinstance(key[2], Decimal)
