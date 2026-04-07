from __future__ import annotations

from datetime import date
from decimal import Decimal

import pytest

from ledgerflow.config import Config
from ledgerflow.utils.exchange_rates import convert_amount_to_ils


def test_convert_amount_to_ils_uses_manual_rate_from_config():
    config = Config()
    config.currency_rates = {"USD": Decimal("3.149")}

    result = convert_amount_to_ils(
        Decimal("20.00"),
        "USD",
        date(2026, 3, 19),
        config,
    )

    assert result == Decimal("62.98")


def test_convert_amount_to_ils_raises_when_manual_rate_missing():
    config = Config()

    with pytest.raises(ValueError, match="USD"):
        convert_amount_to_ils(
            Decimal("20.00"),
            "USD",
            date(2026, 3, 19),
            config,
        )
