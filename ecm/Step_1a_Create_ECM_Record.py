"""Step 1a: create an ECM review worksheet pending manual first-date checks.

This script intentionally leaves otherwise eligible Count (Y/N) cells blank. Users
enter the verified first document date in column M (Remark), then Step 1b finalizes
Count (Y/N) without recreating the worksheet.
"""

from __future__ import annotations

import argparse
import shutil
import sys
from copy import copy
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import Border

from Step_1_Create_ECM_Record import (
    CSV_DIRECTORY,
    DEFAULT_ECM_RECORD,
    DEFAULT_ECM_TEMPLATE,
    DEFAULT_OUTPUT_DIRECTORY,
    DEFAULT_SMM_WORKBOOK,
    REVIEW_HEADERS,
    EcmRecord,
    InputValidationError,
    canonical_project_code,
    find_newest_csv,
    parse_date,
    period_sheet_name,
    read_ecm_records,
    read_readiness,
    get_template_review_worksheet,
)

OUTGOING_REFERENCE_COLUMN = 16
OUTGOING_REFERENCE_HEADER = "Outgoing ref. no. (without revision)"


def outgoing_reference_without_revision(outgoing_reference: str, document_revision: str) -> str:
    """Remove a matching terminal slash/revision suffix from an outgoing reference."""
    reference = outgoing_reference.strip()
    revision = document_revision.strip()
    suffix = f"/{revision}"
    if revision and reference.endswith(suffix):
        return reference[: -len(suffix)]
    return reference


def apply_preliminary_counting_rules(records: list[EcmRecord], readiness: dict[int, bool]) -> None:
    """Apply all Step 1 rules except the manual first-document-date decision."""
    highest_versions: dict[str, int] = {}
    for record in records:
        record.reasons.clear()
        record.count = ""
        highest_versions[record.normalized_name] = max(
            highest_versions.get(record.normalized_name, record.major_version), record.major_version
        )

    candidates_by_name: dict[str, list[EcmRecord]] = {}
    for record in records:
        if record.project_code is None:
            record.reasons.append("unresolved_project_code")
        elif not readiness.get(canonical_project_code(record.project_code), False):
            record.reasons.append("project_not_ready")

        if record.major_version < highest_versions[record.normalized_name]:
            record.reasons.append("superseded_major_version")
        else:
            candidates_by_name.setdefault(record.normalized_name, []).append(record)

    for candidates in candidates_by_name.values():
        if len(candidates) <= 1:
            continue
        canonical = max(
            candidates,
            key=lambda record: (record.modified_on or datetime.min, -record.source_row),
        )
        for record in candidates:
            if record is not canonical:
                record.reasons.append("duplicate_max_major_version")

    for record in records:
        # A blank value explicitly means that the manual M-column date is required.
        record.count = "N" if record.reasons else ""


def validate_review_headers(worksheet: Any) -> None:
    headings = tuple(str(worksheet.cell(1, column).value or "").strip() for column in range(1, 12))
    if headings != REVIEW_HEADERS:
        raise InputValidationError(
            f"Review worksheet headers A:K do not match the required ECM layout: {headings!r}."
        )
    if str(worksheet.cell(1, 13).value or "").strip() != "Remark":
        raise InputValidationError("Review worksheet column M must have the header 'Remark'.")
    outgoing_header = str(worksheet.cell(1, OUTGOING_REFERENCE_COLUMN).value or "").strip()
    if outgoing_header and outgoing_header != OUTGOING_REFERENCE_HEADER:
        raise InputValidationError(
            f"Review worksheet column P must have the header {OUTGOING_REFERENCE_HEADER!r}."
        )


def copy_row_style(worksheet: Any, source_row: int, target_row: int) -> None:
    for column in range(1, worksheet.max_column + 1):
        source = worksheet.cell(source_row, column)
        target = worksheet.cell(target_row, column)
        if source.has_style:
            target._style = copy(source._style)
        if source.number_format:
            target.number_format = source.number_format
        if source.alignment:
            target.alignment = copy(source.alignment)
        if source.fill:
            target.fill = copy(source.fill)


def clear_review_rows(worksheet: Any) -> None:
    for row_number in range(2, worksheet.max_row + 1):
        for column in range(1, worksheet.max_column + 1):
            worksheet.cell(row_number, column).value = None
            worksheet.row_dimensions[row_number].hidden = False


def write_preliminary_review_workbook(
    workbook_path: Path,
    start_date: date,
    end_date: date,
    records: list[EcmRecord],
) -> None:
    """Replace the reporting-period worksheet and leave it ready for manual M entry."""
    workbook = load_workbook(workbook_path)
    try:
        source_worksheet = get_template_review_worksheet(workbook, workbook_path.name)
        validate_review_headers(source_worksheet)
        target_sheet_name = period_sheet_name(start_date, end_date)
        if target_sheet_name in workbook.sheetnames:
            del workbook[target_sheet_name]
        worksheet = workbook.copy_worksheet(source_worksheet)
        worksheet.title = target_sheet_name
        worksheet.cell(1, OUTGOING_REFERENCE_COLUMN).value = OUTGOING_REFERENCE_HEADER
        style_row = 2 if worksheet.max_row >= 2 else None
        clear_review_rows(worksheet)

        ordered_records = sorted(
            records,
            key=lambda record: (
                record.project_code is None,
                record.project_code if record.project_code is not None else 0,
                record.name.casefold(),
                record.major_version,
                record.document_date,
            ),
        )
        highest_versions = {
            record.normalized_name: max(
                candidate.major_version
                for candidate in records
                if candidate.normalized_name == record.normalized_name
            )
            for record in records
        }
        for row_number, record in enumerate(ordered_records, start=2):
            if style_row is not None and row_number != style_row:
                copy_row_style(worksheet, style_row, row_number)
            values = (
                record.project_code,
                record.document_title,
                record.document_date,
                record.name,
                record.document_revision or None,
                record.modified_by,
                record.major_version,
                record.smm_id or None,
                record.count or None,
                None,
                None,
            )
            for column, value in enumerate(values, start=1):
                worksheet.cell(row_number, column).value = value
            outgoing_reference = outgoing_reference_without_revision(
                record.outgoing_reference,
                record.document_revision,
            )
            worksheet.cell(row_number, OUTGOING_REFERENCE_COLUMN).value = outgoing_reference or None
            if "superseded_major_version" in record.reasons:
                worksheet.cell(row_number, 13).value = (
                    f"Duplicate with its version {highest_versions[record.normalized_name]}"
                )
            # Column N is populated with a static value during Step 2, after J/K review.

        for row_number in range(len(ordered_records) + 2, worksheet.max_row + 1):
            for column in range(1, worksheet.max_column + 1):
                worksheet.cell(row_number, column).border = Border()

        for row_number in range(1, worksheet.max_row + 1):
            worksheet.row_dimensions[row_number].hidden = False
        worksheet.auto_filter.ref = f"A1:P{max(1, len(ordered_records) + 1)}"
        worksheet.auto_filter.filterColumn = []
        worksheet.freeze_panes = "A2"
        for sheet in workbook.worksheets:
            sheet.sheet_view.tabSelected = False
        worksheet.sheet_view.tabSelected = True
        workbook.active = workbook.index(worksheet)
        workbook.save(workbook_path)
    finally:
        workbook.close()


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Step 1a: create ECM_Record.xlsx for manual first-date review.")
    parser.add_argument("--start-date", required=True, type=parse_date)
    parser.add_argument("--end-date", required=True, type=parse_date)
    parser.add_argument("--csv", type=Path, help="ECM properties CSV. Defaults to the newest properties*.csv.")
    parser.add_argument("--smm-workbook", type=Path, default=DEFAULT_SMM_WORKBOOK)
    parser.add_argument("--ecm-record", type=Path, default=DEFAULT_ECM_RECORD)
    parser.add_argument("--ecm-template", type=Path, default=DEFAULT_ECM_TEMPLATE)
    parser.add_argument("--backup", type=Path, help="Pre-1a backup path.")
    args = parser.parse_args()
    if args.end_date < args.start_date:
        parser.error("--end-date cannot be earlier than --start-date.")
    return args


def main() -> int:
    args = parse_arguments()
    csv_path = args.csv or find_newest_csv(CSV_DIRECTORY)
    backup_path = args.backup or (
        DEFAULT_OUTPUT_DIRECTORY
        / f"ECM_Record_{args.start_date:%Y-%m-%d}_to_{args.end_date:%Y-%m-%d}_1a.xlsx"
    )
    args.ecm_record.parent.mkdir(parents=True, exist_ok=True)
    if not args.ecm_record.is_file():
        if not args.ecm_template.is_file():
            raise FileNotFoundError(f"ECM review template was not found: {args.ecm_template}")
        shutil.copy2(args.ecm_template, args.ecm_record)
    if backup_path.resolve() == args.ecm_record.resolve():
        raise InputValidationError("The 1a backup path must differ from the ECM_Record.xlsx path.")

    records = read_ecm_records(csv_path, args.start_date, args.end_date)
    readiness = read_readiness(args.smm_workbook)
    apply_preliminary_counting_rules(records, readiness)
    shutil.copy2(args.ecm_record, backup_path)
    write_preliminary_review_workbook(args.ecm_record, args.start_date, args.end_date, records)

    print(f"Input CSV: {csv_path.resolve()}")
    print(f"Readiness workbook: {args.smm_workbook.resolve()}")
    print(f"Updated ECM record: {args.ecm_record.resolve()}")
    print(f"Backup: {backup_path.resolve()}")
    print(f"Rows written: {len(records)}")
    print(f"Pending manual first dates: {sum(record.count == '' for record in records)}")
    print(f"Count N: {sum(record.count == 'N' for record in records)}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (FileNotFoundError, InputValidationError, OSError, ValueError) as error:
        print(f"Error: {error}", file=sys.stderr)
        raise SystemExit(1)
