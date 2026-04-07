from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass

from ledgerflow.config import Config
from ledgerflow.google_sheets import (
    RowUpdate,
    SheetsGateway,
    build_sheets_gateway,
    spreadsheet_url,
)
from ledgerflow.models import Category, Transaction
from ledgerflow.year_writer import (
    CategoryBlock,
    CategoryBlockNotFoundError,
    detect_category_blocks,
    find_first_empty_row,
)


@dataclass(frozen=True)
class WriteResult:
    backup_path: str
    output_path: str
    transactions_written: int
    skipped_count: int


def write_yearly_sheet(
    transactions: list[Transaction],
    spreadsheet_id: str,
    month: int,
    config: Config,
    gateway: SheetsGateway | None = None,
) -> WriteResult:
    sheets = gateway or build_sheets_gateway(config.google_service_account_file)
    resolved_spreadsheet_id, worksheet_name = config.resolved_year_target(month)
    if spreadsheet_id != resolved_spreadsheet_id:
        raise ValueError(
            f"Year spreadsheet ID mismatch. Expected {resolved_spreadsheet_id!r}, got {spreadsheet_id!r}."
        )
    rows = [tuple(row) for row in sheets.read_sheet(spreadsheet_id, worksheet_name)]
    blocks = detect_category_blocks(rows, config)

    block_assignments: dict[CategoryBlock, list[Transaction]] = defaultdict(list)
    for tx in transactions:
        category = tx.category if tx.category is not None else Category.OTHER
        if category in blocks:
            block_assignments[blocks[category]].append(tx)
        elif Category.OTHER in blocks:
            block_assignments[blocks[Category.OTHER]].append(tx)
        else:
            raise CategoryBlockNotFoundError(
                f"Transaction category {category.value!r} has no block in worksheet "
                f"{worksheet_name!r}, and no 'Other' fallback block was detected."
            )

    mutable_rows: list[list[object]] = [list(row) for row in rows]
    updates: list[RowUpdate] = []
    written = 0

    for block, block_transactions in block_assignments.items():
        insert_row = find_first_empty_row([tuple(row) for row in mutable_rows], block)
        for tx in block_transactions:
            _ensure_cell(mutable_rows, insert_row, block.date_col)
            mutable_rows[insert_row][block.merchant_col] = tx.merchant
            mutable_rows[insert_row][block.price_col] = float(tx.amount)
            mutable_rows[insert_row][block.date_col] = tx.date
            updates.append(
                RowUpdate(
                    row=insert_row,
                    start_col=block.merchant_col,
                    values=[
                        tx.merchant,
                        float(tx.amount),
                        tx.date,
                    ],
                )
            )
            insert_row += 1
            written += 1

    sheets.batch_update_rows(spreadsheet_id, worksheet_name, updates)
    return WriteResult(
        backup_path="Google Sheets write: no local backup created",
        output_path=spreadsheet_url(spreadsheet_id),
        transactions_written=written,
        skipped_count=0,
    )


def _ensure_cell(rows: list[list[object]], row_idx: int, col_idx: int) -> None:
    while len(rows) <= row_idx:
        rows.append([])
    row = rows[row_idx]
    if len(row) <= col_idx:
        row.extend([""] * (col_idx + 1 - len(row)))
