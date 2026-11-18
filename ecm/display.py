"""Static display-value support for the ECM review worksheet Display column."""

from __future__ import annotations

from typing import Any

from openpyxl import load_workbook

DISPLAY_COLUMN = 14


def display_value(count: Any, use_smm: Any, use_kms: Any) -> str:
    """Return the static Display text for one ECM row."""
    flags = tuple(str(value or "").strip().upper() for value in (count, use_smm, use_kms))
    if flags == ("Y", "Y", "Y"):
        return "via KMS"
    if flags == ("Y", "Y", "N"):
        return "w/o KMS"
    return ""


def populate_display_values(worksheet: Any, data_rows: list[int]) -> None:
    """Set static column-N values for data rows without changing other columns."""
    for row_number in data_rows:
        worksheet.cell(row_number, DISPLAY_COLUMN).value = display_value(
            worksheet.cell(row_number, 9).value,
            worksheet.cell(row_number, 10).value,
            worksheet.cell(row_number, 11).value,
        )


def write_display_values(workbook_path: Any, worksheet_name: str) -> None:
    """Write static Display values to one workbook worksheet."""
    workbook = load_workbook(workbook_path)
    try:
        worksheet = workbook[worksheet_name]
        data_rows = [
            row_number
            for row_number in range(2, worksheet.max_row + 1)
            if any(worksheet.cell(row_number, column).value is not None for column in range(1, 9))
        ]
        populate_display_values(worksheet, data_rows)
        workbook.save(workbook_path)
    finally:
        workbook.close()
