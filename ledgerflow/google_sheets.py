from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Protocol


GOOGLE_SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets"


class GoogleSheetsConfigError(ValueError):
    """Google Sheets credentials or configuration are missing."""


class WorksheetNotFoundError(ValueError):
    """The target worksheet does not exist in the spreadsheet."""


@dataclass(frozen=True)
class RowUpdate:
    row: int
    start_col: int
    values: list[object]


class SheetsGateway(Protocol):
    def list_sheet_names(self, spreadsheet_id: str) -> list[str]:
        ...

    def read_sheet(self, spreadsheet_id: str, worksheet_name: str) -> list[list[object]]:
        ...

    def batch_update_rows(
        self,
        spreadsheet_id: str,
        worksheet_name: str,
        updates: list[RowUpdate],
    ) -> None:
        ...

    def append_rows(
        self,
        spreadsheet_id: str,
        worksheet_name: str,
        rows: list[list[object]],
    ) -> None:
        ...

    def read_range(
        self,
        spreadsheet_id: str,
        worksheet_name: str,
        a1_range: str,
    ) -> list[list[object]]:
        ...

    def update_range(
        self,
        spreadsheet_id: str,
        worksheet_name: str,
        a1_range: str,
        values: list[list[object]],
    ) -> None:
        ...


def spreadsheet_url(spreadsheet_id: str) -> str:
    return f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"


def build_sheets_gateway(
    service_account_file: Path | None = None,
) -> SheetsGateway:
    resolved_creds = service_account_file or _service_account_from_env()
    if resolved_creds is None:
        raise GoogleSheetsConfigError(
            "Google Sheets credentials are not configured. "
            "Set GOOGLE_SERVICE_ACCOUNT_FILE or configure google_service_account_file."
        )

    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
    except ImportError as exc:
        raise GoogleSheetsConfigError(
            "Google Sheets dependencies are not installed. "
            "Install google-api-python-client and google-auth."
        ) from exc

    credentials = Credentials.from_service_account_file(
        str(resolved_creds),
        scopes=[GOOGLE_SHEETS_SCOPE],
    )
    service = build("sheets", "v4", credentials=credentials, cache_discovery=False)
    return GoogleSheetsGateway(service)


class GoogleSheetsGateway:
    def __init__(self, service: object) -> None:
        self._service = service

    def list_sheet_names(self, spreadsheet_id: str) -> list[str]:
        response = (
            self._service.spreadsheets()
            .get(spreadsheetId=spreadsheet_id, fields="sheets.properties.title")
            .execute()
        )
        names: list[str] = []
        for sheet in response.get("sheets", []):
            props = sheet.get("properties", {})
            title = props.get("title")
            if isinstance(title, str):
                names.append(title)
        return names

    def read_sheet(self, spreadsheet_id: str, worksheet_name: str) -> list[list[object]]:
        self._ensure_sheet_exists(spreadsheet_id, worksheet_name)
        response = (
            self._service.spreadsheets()
            .values()
            .get(
                spreadsheetId=spreadsheet_id,
                range=_sheet_range(worksheet_name),
                valueRenderOption="UNFORMATTED_VALUE",
            )
            .execute()
        )
        return [list(row) for row in response.get("values", [])]

    def read_range(
        self,
        spreadsheet_id: str,
        worksheet_name: str,
        a1_range: str,
    ) -> list[list[object]]:
        self._ensure_sheet_exists(spreadsheet_id, worksheet_name)
        response = (
            self._service.spreadsheets()
            .values()
            .get(
                spreadsheetId=spreadsheet_id,
                range=f"{_sheet_range(worksheet_name)}!{a1_range}",
                valueRenderOption="UNFORMATTED_VALUE",
            )
            .execute()
        )
        return [list(row) for row in response.get("values", [])]

    def batch_update_rows(
        self,
        spreadsheet_id: str,
        worksheet_name: str,
        updates: list[RowUpdate],
    ) -> None:
        if not updates:
            return
        self._ensure_sheet_exists(spreadsheet_id, worksheet_name)
        data = []
        for update in updates:
            start_cell = _a1(update.row, update.start_col)
            end_cell = _a1(update.row, update.start_col + len(update.values) - 1)
            data.append(
                {
                    "range": f"{_sheet_range(worksheet_name)}!{start_cell}:{end_cell}",
                    "values": [[_serialize_sheet_value(value) for value in update.values]],
                }
            )
        (
            self._service.spreadsheets()
            .values()
            .batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"valueInputOption": "USER_ENTERED", "data": data},
            )
            .execute()
        )

    def append_rows(
        self,
        spreadsheet_id: str,
        worksheet_name: str,
        rows: list[list[object]],
    ) -> None:
        if not rows:
            return
        self._ensure_sheet_exists(spreadsheet_id, worksheet_name)
        (
            self._service.spreadsheets()
            .values()
            .append(
                spreadsheetId=spreadsheet_id,
                range=_sheet_range(worksheet_name),
                valueInputOption="USER_ENTERED",
                insertDataOption="INSERT_ROWS",
                body={
                    "values": [
                        [_serialize_sheet_value(value) for value in row]
                        for row in rows
                    ]
                },
            )
            .execute()
        )

    def update_range(
        self,
        spreadsheet_id: str,
        worksheet_name: str,
        a1_range: str,
        values: list[list[object]],
    ) -> None:
        self._ensure_sheet_exists(spreadsheet_id, worksheet_name)
        (
            self._service.spreadsheets()
            .values()
            .update(
                spreadsheetId=spreadsheet_id,
                range=f"{_sheet_range(worksheet_name)}!{a1_range}",
                valueInputOption="USER_ENTERED",
                body={
                    "values": [
                        [_serialize_sheet_value(value) for value in row]
                        for row in values
                    ]
                },
            )
            .execute()
        )

    def _ensure_sheet_exists(self, spreadsheet_id: str, worksheet_name: str) -> None:
        if worksheet_name not in self.list_sheet_names(spreadsheet_id):
            raise WorksheetNotFoundError(
                f"Worksheet {worksheet_name!r} was not found in spreadsheet {spreadsheet_id!r}."
            )


def _service_account_from_env() -> Path | None:
    raw = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "").strip()
    return Path(raw) if raw else None


def _sheet_range(worksheet_name: str) -> str:
    escaped = worksheet_name.replace("'", "''")
    return f"'{escaped}'"


def _serialize_sheet_value(value: object) -> object:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return value


def _a1(row: int, col: int) -> str:
    return f"{_col_letters(col)}{row + 1}"


def _col_letters(col: int) -> str:
    current = col + 1
    letters = ""
    while current > 0:
        current, remainder = divmod(current - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters
