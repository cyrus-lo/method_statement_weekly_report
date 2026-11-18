"""Step 1b: finalize Count (Y/N) after revision-gated first-date review in M.

Only Count=N rows are skipped. Blank and Y Count rows with a non-empty Document
revision require a first date in M. Rows without a Document revision become Y
without an M date. A successful output has no blank Count cells.
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
from copy import copy
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.worksheet.filters import FilterColumn, Filters

from Step_1_Create_ECM_Record import (
    DEFAULT_ECM_RECORD,
    DEFAULT_OUTPUT_DIRECTORY,
    FIRST_DATE_CUTOFF,
    MISSING_SMM_ID_FILL,
    REVIEW_HEADERS,
    InputValidationError,
    apply_count_filter_in_excel,
    parse_date,
    period_sheet_name,
)
from Step_1a_Create_ECM_Record import (
    OUTGOING_REFERENCE_COLUMN,
    OUTGOING_REFERENCE_HEADER,
)

REMARK_COLUMN = 13
COUNT_COLUMN = 9
SMM_ID_COLUMN = 8
DOCUMENT_REVISION_COLUMN = 5
ISO_DATE_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}$")
FIRST_VERSION_REMARK_PATTERN = re.compile(
    r"^1st version is submitted on (\d{1,2} [A-Za-z]{3} \d{4})$"
)


def validate_review_headers(worksheet: Any) -> None:
    headings = tuple(str(worksheet.cell(1, column).value or "").strip() for column in range(1, 12))
    if headings != REVIEW_HEADERS:
        raise InputValidationError(
            f"Review worksheet headers A:K do not match the required ECM layout: {headings!r}."
        )
    if str(worksheet.cell(1, REMARK_COLUMN).value or "").strip() != "Remark":
        raise InputValidationError("Review worksheet column M must have the header 'Remark'.")
    if (
        str(worksheet.cell(1, OUTGOING_REFERENCE_COLUMN).value or "").strip()
        != OUTGOING_REFERENCE_HEADER
    ):
        raise InputValidationError(
            f"Review worksheet column P must have the header {OUTGOING_REFERENCE_HEADER!r}."
        )


def parse_manual_first_date(value: Any, row_number: int) -> date:
    """Accept an Excel date, ISO date, or the saved first-version remark format."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        text = value.strip()
        if ISO_DATE_PATTERN.fullmatch(text):
            try:
                return date.fromisoformat(text)
            except ValueError as error:
                raise InputValidationError(
                    f"Worksheet row {row_number}: M (Remark) must be an Excel date, YYYY-MM-DD, "
                    f"or a saved first-version remark; got {value!r}."
                ) from error
        saved_date_match = FIRST_VERSION_REMARK_PATTERN.fullmatch(text)
        if saved_date_match:
            try:
                return datetime.strptime(saved_date_match.group(1), "%d %b %Y").date()
            except ValueError as error:
                raise InputValidationError(
                    f"Worksheet row {row_number}: M (Remark) has an invalid saved first-version date; "
                    f"got {value!r}."
                ) from error
        if text:
            raise InputValidationError(
                f"Worksheet row {row_number}: M (Remark) must be an Excel date, YYYY-MM-DD, "
                f"or a saved first-version remark; got {value!r}."
            )
    raise InputValidationError(
        f"Worksheet row {row_number}: M (Remark) requires a verified first document date for a pending Count row."
    )


def first_version_remark(first_date: date) -> str:
    """Return the standard saved remark for a verified first-version date."""
    return f"1st version is submitted on {first_date.day} {first_date:%b %Y}"


def populated_data_rows(worksheet: Any) -> list[int]:
    """Return non-header rows containing ECM source data, without trusting max_row alone."""
    return [
        row_number
        for row_number in range(2, worksheet.max_row + 1)
        if any(worksheet.cell(row_number, column).value is not None for column in range(1, 9))
    ]


def finalize_pending_rows(
    workbook_path: Path,
    worksheet_name: str,
    backup_path: Path,
) -> tuple[int, int]:
    """Validate all non-N rows, back up, then save Count decisions."""
    if not workbook_path.is_file():
        raise FileNotFoundError(f"ECM review workbook was not found: {workbook_path}")
    if backup_path.resolve() == workbook_path.resolve():
        raise InputValidationError("The 1b backup path must differ from the ECM_Record.xlsx path.")
    workbook = load_workbook(workbook_path)
    try:
        if worksheet_name not in workbook.sheetnames:
            raise InputValidationError(f"Review worksheet {worksheet_name!r} was not found in {workbook_path.name}.")
        worksheet = workbook[worksheet_name]
        validate_review_headers(worksheet)
        data_rows = populated_data_rows(worksheet)

        count_decisions: list[tuple[int, str, date | None]] = []
        for row_number in data_rows:
            count_value = str(worksheet.cell(row_number, COUNT_COLUMN).value or "").strip().upper()
            if count_value not in {"", "Y", "N"}:
                raise InputValidationError(
                    f"Worksheet row {row_number}: Count (Y/N) must be blank, Y, or N; got {count_value!r}."
                )
            if count_value == "N":
                continue

            document_revision = str(
                worksheet.cell(row_number, DOCUMENT_REVISION_COLUMN).value or ""
            ).strip()
            if not document_revision:
                count_decisions.append((row_number, "Y", None))
                continue

            first_date = parse_manual_first_date(worksheet.cell(row_number, REMARK_COLUMN).value, row_number)
            final_count = "N" if first_date <= FIRST_DATE_CUTOFF else "Y"
            count_decisions.append((row_number, final_count, first_date))

        # Every input is validated before the persistent workbook or backup is changed.
        backup_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(workbook_path, backup_path)
        for row_number, final_count, first_date in count_decisions:
            worksheet.cell(row_number, COUNT_COLUMN).value = final_count
            if first_date is not None:
                worksheet.cell(row_number, REMARK_COLUMN).value = first_version_remark(first_date)
        for row_number in data_rows:
            if not str(worksheet.cell(row_number, COUNT_COLUMN).value or "").strip():
                worksheet.cell(row_number, COUNT_COLUMN).value = "Y"
            if str(worksheet.cell(row_number, SMM_ID_COLUMN).value or "").strip():
                continue
            for column in range(1, worksheet.max_column + 1):
                worksheet.cell(row_number, column).fill = copy(MISSING_SMM_ID_FILL)

        last_row = max(data_rows, default=1)
        worksheet.auto_filter.ref = f"A1:P{last_row}"
        worksheet.auto_filter.filterColumn = [
            FilterColumn(colId=COUNT_COLUMN - 1, filters=Filters(filter=["Y"]))
        ]
        for sheet in workbook.worksheets:
            sheet.sheet_view.tabSelected = False
        worksheet.sheet_view.tabSelected = True
        workbook.active = workbook.index(worksheet)
        temporary_workbook_path = workbook_path.with_name(
            f".{workbook_path.name}.step1b"
        )
        try:
            workbook.save(temporary_workbook_path)
            os.replace(temporary_workbook_path, workbook_path)
        finally:
            if temporary_workbook_path.exists():
                temporary_workbook_path.unlink()
        return (
            sum(final_count == "Y" for _, final_count, _ in count_decisions),
            sum(final_count == "N" for _, final_count, _ in count_decisions),
        )
    finally:
        workbook.close()


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Step 1b: finalize ECM_Record.xlsx after manual M-column review.")
    parser.add_argument("--start-date", required=True, type=parse_date)
    parser.add_argument("--end-date", required=True, type=parse_date)
    parser.add_argument("--ecm-record", type=Path, default=DEFAULT_ECM_RECORD)
    parser.add_argument("--ecm-sheet", help="ECM review sheet. Defaults to the sheet name derived from the period.")
    parser.add_argument("--backup", type=Path, help="Pre-1b backup path.")
    parser.add_argument(
        "--post-1b-backup",
        type=Path,
        help="Post-1b backup created after all Step 1b Excel operations.",
    )
    args = parser.parse_args()
    if args.end_date < args.start_date:
        parser.error("--end-date cannot be earlier than --start-date.")
    return args


def main() -> int:
    args = parse_arguments()
    worksheet_name = args.ecm_sheet or period_sheet_name(args.start_date, args.end_date)
    backup_path = args.backup or (
        DEFAULT_OUTPUT_DIRECTORY
        / f"ECM_Record_{args.start_date:%Y-%m-%d}_to_{args.end_date:%Y-%m-%d}_1b.xlsx"
    )
    post_1b_backup_path = args.post_1b_backup or (
        DEFAULT_OUTPUT_DIRECTORY
        / f"ECM_Record_{args.start_date:%Y-%m-%d}_to_{args.end_date:%Y-%m-%d}_1b_after.xlsx"
    )
    count_y, count_n = finalize_pending_rows(
        args.ecm_record,
        worksheet_name,
        backup_path,
    )
    apply_count_filter_in_excel(args.ecm_record, worksheet_name)
    post_1b_backup_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(args.ecm_record, post_1b_backup_path)

    print(f"Updated ECM record: {args.ecm_record.resolve()}")
    print(f"Backup: {backup_path.resolve()}")
    print(f"Post-1b backup: {post_1b_backup_path.resolve()}")
    print(f"Processed non-N Count Y: {count_y}")
    print(f"Processed non-N Count N: {count_n}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (FileNotFoundError, InputValidationError, OSError, RuntimeError, ValueError) as error:
        print(f"Error: {error}", file=sys.stderr)
        raise SystemExit(1)
