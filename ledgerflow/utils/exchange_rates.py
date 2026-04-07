"""
Manual exchange-rate lookup for normalizing foreign charges to ILS.
"""
from __future__ import annotations

from datetime import date
from decimal import Decimal, ROUND_HALF_UP

from ledgerflow.config import Config


def get_rate_to_ils(
    currency_code: str,
    conversion_date: date,
    config: Config,
) -> Decimal:
    del conversion_date  # Manual rates are not date-sensitive.

    normalized = currency_code.upper()
    if normalized == "ILS":
        return Decimal("1")

    rate = config.currency_rates.get(normalized)
    if rate is None:
        raise ValueError(
            f"No manual currency rate is configured for {normalized}."
        )
    return Decimal(rate)


def convert_amount_to_ils(
    amount: Decimal,
    currency_code: str,
    conversion_date: date,
    config: Config,
) -> Decimal:
    rate = get_rate_to_ils(currency_code, conversion_date, config)
    return (amount * rate).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
