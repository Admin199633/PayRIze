from __future__ import annotations

import json
from decimal import Decimal
from pathlib import Path

from ledgerflow.app import (
    json_dumps_config,
    load_default_config,
    resolve_spreadsheet_id,
    save_config,
    validate_currency_rate_inputs,
)
from ledgerflow.config import Config


def test_load_default_config_returns_none_when_missing(tmp_path: Path):
    assert load_default_config(tmp_path / "config.json") is None


def test_load_default_config_reads_seed_and_year_spreadsheet_ids(tmp_path: Path):
    config_path = tmp_path / "config.json"
    config_path.write_text(
        json.dumps(
            {
                "seed": {
                    "spreadsheet_id": "seed-sheet-id",
                    "worksheet_name": "Seed",
                },
                "year": {
                    "spreadsheet_id": "year-sheet-id",
                    "month_sheets": {"3": "March"},
                },
            }
        ),
        encoding="utf-8",
    )

    config = load_default_config(config_path)

    assert config is not None
    assert config.seed_spreadsheet_id == "seed-sheet-id"
    assert config.year_spreadsheet_id == "year-sheet-id"


def test_load_default_config_reads_currency_rates(tmp_path: Path):
    config_path = tmp_path / "config.json"
    config_path.write_text(
        json.dumps(
            {
                "currency_rates": {
                    "USD": 3.149,
                    "EUR": 3.92,
                    "GBP": 4.61,
                }
            }
        ),
        encoding="utf-8",
    )

    config = load_default_config(config_path)

    assert config is not None
    assert config.currency_rates == {
        "USD": Decimal("3.149"),
        "EUR": Decimal("3.92"),
        "GBP": Decimal("4.61"),
    }


def test_load_default_config_resolves_credentials_relative_to_config_file(tmp_path: Path):
    config_path = tmp_path / "config.json"
    config_path.write_text(
        json.dumps(
            {
                "google_service_account_file": "credentials.json",
            }
        ),
        encoding="utf-8",
    )

    config = load_default_config(config_path)

    assert config is not None
    assert config.google_service_account_file == tmp_path / "credentials.json"


def test_resolve_spreadsheet_id_prefers_manual_value():
    assert resolve_spreadsheet_id("manual-id", "config-id") == "manual-id"


def test_resolve_spreadsheet_id_falls_back_to_config_value():
    assert resolve_spreadsheet_id("", "config-id") == "config-id"


def test_resolve_spreadsheet_id_returns_empty_when_both_missing():
    assert resolve_spreadsheet_id("   ", None) == ""


def test_validate_currency_rate_inputs_accepts_positive_numeric_values():
    rates = validate_currency_rate_inputs({
        "USD": "3.149",
        "EUR": "3.92",
        "GBP": "4.61",
    })

    assert rates == {
        "USD": Decimal("3.149"),
        "EUR": Decimal("3.92"),
        "GBP": Decimal("4.61"),
    }


def test_validate_currency_rate_inputs_rejects_invalid_value():
    try:
        validate_currency_rate_inputs({
            "USD": "abc",
            "EUR": "3.92",
            "GBP": "4.61",
        })
    except ValueError as exc:
        assert "USD" in str(exc)
    else:
        raise AssertionError("Expected ValueError for invalid currency input")


def test_save_config_persists_currency_rates(tmp_path: Path):
    config = Config()
    config.seed_spreadsheet_id = "seed-sheet-id"
    config.year_spreadsheet_id = "year-sheet-id"
    config.year_worksheet_names = {3: "March"}
    config.currency_rates = {
        "USD": Decimal("3.149"),
        "EUR": Decimal("3.92"),
        "GBP": Decimal("4.61"),
    }
    path = tmp_path / "config.json"

    save_config(path, config)

    saved = json.loads(path.read_text(encoding="utf-8"))
    assert saved["currency_rates"] == {
        "USD": 3.149,
        "EUR": 3.92,
        "GBP": 4.61,
    }


def test_json_dumps_config_includes_currency_rates():
    config = Config()
    config.currency_rates = {
        "USD": Decimal("3.149"),
        "EUR": Decimal("3.92"),
        "GBP": Decimal("4.61"),
    }

    payload = json.loads(json_dumps_config(config))

    assert payload["currency_rates"] == {
        "USD": 3.149,
        "EUR": 3.92,
        "GBP": 4.61,
    }
