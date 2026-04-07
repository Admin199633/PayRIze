from __future__ import annotations

from datetime import date
from decimal import Decimal
from pathlib import Path

from ledgerflow.config import Config
from ledgerflow.google_year_writer import write_yearly_sheet
from ledgerflow.models import Category, Transaction


class FakeSheetsGateway:
    def __init__(self, sheets: dict[str, dict[str, list[list[object]]]]) -> None:
        self.sheets = sheets

    def list_sheet_names(self, spreadsheet_id: str) -> list[str]:
        return list(self.sheets[spreadsheet_id].keys())

    def read_sheet(self, spreadsheet_id: str, worksheet_name: str) -> list[list[object]]:
        return [list(row) for row in self.sheets[spreadsheet_id][worksheet_name]]

    def batch_update_rows(self, spreadsheet_id: str, worksheet_name: str, updates) -> None:
        sheet = self.sheets[spreadsheet_id][worksheet_name]
        for update in updates:
            while len(sheet) <= update.row:
                sheet.append([])
            row = sheet[update.row]
            needed = update.start_col + len(update.values)
            if len(row) < needed:
                row.extend([""] * (needed - len(row)))
            row[update.start_col:needed] = list(update.values)

    def append_rows(self, spreadsheet_id: str, worksheet_name: str, rows) -> None:
        raise AssertionError("Year writer should not append whole rows directly.")


def tx(
    merchant: str,
    amount: str,
    d: str = "2024-03-15",
    category: Category = Category.OTHER,
) -> Transaction:
    return Transaction(
        date=date.fromisoformat(d),
        merchant=merchant,
        normalized_merchant=merchant.lower(),
        amount=Decimal(amount),
        source_file=Path("bank.xlsx"),
        sheet_name="Sheet1",
        category=category,
    )


def test_write_yearly_sheet_writes_only_target_cells():
    gateway = FakeSheetsGateway(
        {
            "year-sheet": {
                "March": [
                    ["Other", "Price", "Date"],
                    [],
                    [],
                    [],
                    [],
                    [],
                    ["", "", "", "", "", "keep-me", 111, "2024-01-01"],
                ]
            }
        }
    )

    result = write_yearly_sheet(
        [tx("Cinema", "35.00", category=Category.OTHER)],
        "year-sheet",
        3,
        Config(),
        gateway=gateway,
    )

    rows = gateway.sheets["year-sheet"]["March"]
    assert rows[6][0] == "Cinema"
    assert rows[6][1] == 35.0
    assert str(rows[6][2]) == "2024-03-15"
    assert rows[6][5] == "keep-me"
    assert rows[6][6] == 111
    assert rows[6][7] == "2024-01-01"
    assert result.transactions_written == 1


def test_write_yearly_sheet_leaves_unrelated_content_untouched():
    gateway = FakeSheetsGateway(
        {
            "year-sheet": {
                "March": [
                    ["Other", "Price", "Date", None, None, "Business", "Price", "Date", None],
                    [],
                    [],
                    [],
                    [],
                    [],
                    [None, None, None, None, None, "do-not-touch", 99.0, "2024-01-01", "right-side-note"],
                ]
            }
        }
    )

    write_yearly_sheet(
        [tx("Grocer", "230.00", category=Category.OTHER)],
        "year-sheet",
        3,
        Config(),
        gateway=gateway,
    )

    row = gateway.sheets["year-sheet"]["March"][6]
    assert row[0] == "Grocer"
    assert row[1] == 230.0
    assert str(row[2]) == "2024-03-15"
    assert row[5] == "do-not-touch"
    assert row[6] == 99.0
    assert row[7] == "2024-01-01"
    assert row[8] == "right-side-note"
