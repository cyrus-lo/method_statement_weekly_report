"""Update the SMM working workbook with a period Summary from a reviewed ECM record."""

from __future__ import annotations

import argparse
import re
import shutil
import sys
from collections import defaultdict
from copy import copy
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook

from display import write_display_values

ROOT = Path(__file__).resolve().parent
DEFAULT_OUTPUT_DIRECTORY = ROOT / "output"
DEFAULT_ECM_RECORD = DEFAULT_OUTPUT_DIRECTORY / "ECM_Record.xlsx"
DEFAULT_SMM_TEMPLATE = ROOT / "SMM_KMS_PM_Master_Document_template.xlsx"
DEFAULT_SMM_RECORD = DEFAULT_OUTPUT_DIRECTORY / "SMM_KMS_PM_Master_Document.xlsx"
DEFAULT_SUMMARY_TEMPLATE_SHEET = "(OK) 1 Jul - 12 Jul"
READINESS_SHEET_NAME = "Project Realiness Chris-In Use"
SUMMARY_HEADER_ROWS = 4
SUMMARY_FIRST_DATA_ROW = 5
SUMMARY_FIGURE_COLUMNS = (5, 6, 7, 8)
SUMMARY_FONT_BLACK = "FF000000"
SUMMARY_FONT_RED = "FFFF0000"
# These internal openpyxl widths produce the requested Excel UI widths for the
# workbook/font combination, following the Django workflow calibration.
SUMMARY_COLUMN_WIDTHS = {
    "B": 93.1875,  # Excel UI: 92.55
    "C": 16.46375,  # Excel UI: 15.82
    "E": 20.46375,  # Excel UI: 19.82
    "F": 20.46375,  # Excel UI: 19.82
    "G": 20.46375,  # Excel UI: 19.82
    "H": 27.46375,  # Excel UI: 26.82
}

ECM_HEADERS = {
    1: "Project Code",
    8: "SMM ID",
    9: "Count (Y/N)",
    10: "Use SMM (Y/N)",
    11: "Use KMS (Y/N)",
    13: "Remark",
}
PROJECT_CODE_PATTERN = re.compile(r"^\s*(\d+)\b")
MATTEX_ADDED_REMARKS = (
    "added by mattex",
    "add by chris's team",
)
CANONICAL_PROJECT_DISPLAY = {
    1016: "1016 - 1/WSD/19 (1007)",
}


class InputValidationError(ValueError):
    """Raised when an input workbook does not satisfy the Step 2 contract."""


@dataclass
class ProjectMetrics:
    draft_via_kms: int = 0
    draft_not_via_kms: int = 0
    not_using_smm: int = 0
    mattex_added: int = 0

    @property
    def raw_outbound_ecm(self) -> int:
        return self.draft_via_kms + self.draft_not_via_kms + self.not_using_smm

    @property
    def outbound_ecm(self) -> int:
        return self.raw_outbound_ecm - self.mattex_added


@dataclass(frozen=True)
class ProjectMetadata:
    project_code: int
    project: Any
    business_unit: Any


def parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError as error:
        raise argparse.ArgumentTypeError("Dates must use YYYY-MM-DD format.") from error


def normalise_text(value: Any) -> str:
    return str(value or "").strip()


def is_mattex_added(value: Any) -> bool:
    """Return whether Remark contains an approved Mattex-added marker."""
    remark = " ".join(
        normalise_text(value).casefold().replace("’", "'").replace("‘", "'").split()
    )
    return any(marker in remark for marker in MATTEX_ADDED_REMARKS)


def canonical_project_code(project_code: int | None) -> int | None:
    """Return the readiness-sheet code used for project-code aliases."""
    return 1016 if project_code == 1007 else project_code


def project_display_value(project: ProjectMetadata) -> Any:
    """Return the required stable display label for a canonical project."""
    return CANONICAL_PROJECT_DISPLAY.get(project.project_code, project.project)


def normalise_flag(value: Any, *, field: str, row_number: int, allow_blank: bool) -> str:
    text = normalise_text(value).upper()
    allowed = {"Y", "N"} | ({""} if allow_blank else set())
    if text not in allowed:
        expected = "Y, N, or blank" if allow_blank else "Y or N"
        raise InputValidationError(
            f"ECM row {row_number}: {field} must be {expected}; got {value!r}."
        )
    return text


def normalise_project_code(value: Any) -> int | None:
    if value is None:
        return None
    if isinstance(value, int):
        return canonical_project_code(value)
    if isinstance(value, float) and value.is_integer():
        return canonical_project_code(int(value))
    match = PROJECT_CODE_PATTERN.match(normalise_text(value))
    return canonical_project_code(int(match.group(1))) if match else None


def period_label(start_date: date, end_date: date) -> str:
    if start_date.month == end_date.month and start_date.year == end_date.year:
        return f"{start_date.day} {start_date:%b} - {end_date.day} {end_date:%b}"
    return f"{start_date.day} {start_date:%b} - {end_date.day} {end_date:%b}"


def ecm_sheet_name(start_date: date, end_date: date) -> str:
    if start_date.month == end_date.month and start_date.year == end_date.year:
        return f"ECM_{start_date.day}-{end_date.day} {end_date:%b} (WIP)"
    return f"ECM_{start_date.day} {start_date:%b}-{end_date.day} {end_date:%b} (WIP)"


def summary_sheet_name(start_date: date, end_date: date) -> str:
    return f"(WIP) {period_label(start_date, end_date)}"


def validate_ecm_headers(worksheet: Any) -> None:
    mismatches = []
    for column, expected in ECM_HEADERS.items():
        actual = normalise_text(worksheet.cell(1, column).value)
        if actual != expected:
            mismatches.append(f"{worksheet.cell(1, column).coordinate}: expected {expected!r}, got {actual!r}")
    if mismatches:
        raise InputValidationError("ECM review headers do not match: " + "; ".join(mismatches))


def aggregate_ecm_record(ecm_record_path: Path, worksheet_name: str) -> dict[int, ProjectMetrics]:
    if not ecm_record_path.is_file():
        raise FileNotFoundError(f"ECM record workbook was not found: {ecm_record_path}")

    workbook = load_workbook(ecm_record_path, read_only=True, data_only=True)
    try:
        if worksheet_name not in workbook.sheetnames:
            raise InputValidationError(
                f"ECM worksheet {worksheet_name!r} was not found in {ecm_record_path.name}."
            )
        worksheet = workbook[worksheet_name]
        validate_ecm_headers(worksheet)
        metrics: dict[int, ProjectMetrics] = defaultdict(ProjectMetrics)

        for row_number in range(2, worksheet.max_row + 1):
            values = [worksheet.cell(row_number, column).value for column in ECM_HEADERS]
            if all(value in (None, "") for value in values):
                continue

            count_flag = normalise_flag(
                worksheet.cell(row_number, 9).value,
                field="Count (Y/N)",
                row_number=row_number,
                allow_blank=False,
            )
            if count_flag == "N":
                continue

            use_smm = normalise_flag(
                worksheet.cell(row_number, 10).value,
                field="Use SMM (Y/N)",
                row_number=row_number,
                allow_blank=False,
            )
            use_kms = normalise_flag(
                worksheet.cell(row_number, 11).value,
                field="Use KMS (Y/N)",
                row_number=row_number,
                allow_blank=use_smm == "N",
            )

            project_code = normalise_project_code(worksheet.cell(row_number, 1).value)
            if project_code is None:
                raise InputValidationError(
                    f"ECM row {row_number}: Project Code is required when Count (Y/N) is Y."
                )
            if use_smm != "Y":
                if use_kms == "Y":
                    raise InputValidationError(
                        f"ECM row {row_number}: Use KMS cannot be Y when Use SMM is not Y."
                    )
            project = metrics[project_code]
            mattex_added = is_mattex_added(worksheet.cell(row_number, 13).value)
            if use_smm == "Y" and use_kms == "Y":
                project.draft_via_kms += 1
            elif use_smm == "Y" and use_kms == "N":
                project.draft_not_via_kms += 1
            else:
                if mattex_added:
                    raise InputValidationError(
                        f"ECM row {row_number}: a Mattex-added Remark cannot be used when "
                        "Use SMM is N."
                    )
                project.not_using_smm += 1
                continue
            if mattex_added:
                project.mattex_added += 1

        return dict(metrics)
    finally:
        workbook.close()


def read_project_metadata(workbook: Any) -> dict[int, ProjectMetadata]:
    if READINESS_SHEET_NAME not in workbook.sheetnames:
        raise InputValidationError(
            f"Readiness sheet {READINESS_SHEET_NAME!r} was not found in the SMM workbook."
        )
    worksheet = workbook[READINESS_SHEET_NAME]
    metadata: dict[int, ProjectMetadata] = {}
    for row_number in range(2, worksheet.max_row + 1):
        project_value = worksheet.cell(row_number, 3).value
        project_code = normalise_project_code(project_value)
        if project_code is None:
            continue
        if project_code in metadata:
            raise InputValidationError(
                f"Duplicate project code {project_code} in readiness sheet row {row_number}."
            )
        metadata[project_code] = ProjectMetadata(
            project_code=project_code,
            project=project_value,
            business_unit=worksheet.cell(row_number, 4).value,
        )
    if not metadata:
        raise InputValidationError("No project codes were found in the readiness sheet.")
    return metadata


def validate_summary_layout(worksheet: Any) -> None:
    expected = {
        (2, 2): "Project",
        (2, 3): "BU",
        (2, 4): "SMM (Y/N)",
        (3, 5): "A. Draft via KMS",
        (3, 6): "B. Draft NOT via KMS",
        (3, 7): "Total No. (A + B) Approved",
    }
    mismatches = []
    for (row, column), expected_heading in expected.items():
        actual = " ".join(normalise_text(worksheet.cell(row, column).value).split())
        if actual != expected_heading:
            mismatches.append(
                f"{worksheet.cell(row, column).coordinate}: expected {expected_heading!r}, got {actual!r}"
            )
    outbound_header = " ".join(normalise_text(worksheet.cell(2, 8).value).split())
    if "Record (Outbound) in ECM" not in outbound_header:
        mismatches.append(f"H2: expected an outbound ECM heading, got {outbound_header!r}")
    if mismatches:
        raise InputValidationError("Summary layout does not match: " + "; ".join(mismatches))


def capture_row_styles(worksheet: Any, row_number: int) -> list[Any]:
    return [copy(worksheet.cell(row_number, column)._style) for column in range(1, 9)]


def apply_row_styles(worksheet: Any, row_number: int, styles: Iterable[Any]) -> None:
    for column, style in enumerate(styles, start=1):
        worksheet.cell(row_number, column)._style = copy(style)


def apply_summary_column_formatting(worksheet: Any) -> None:
    for column, width in SUMMARY_COLUMN_WIDTHS.items():
        worksheet.column_dimensions[column].width = width
    worksheet.column_dimensions["D"].hidden = True
    worksheet.sheet_view.zoomScale = 85


def set_cell_font_color(cell: Any, color: str) -> None:
    font = copy(cell.font)
    font.color = color
    cell.font = font


def normalize_summary_figure_font_colors(worksheet: Any) -> None:
    for row_number in range(SUMMARY_FIRST_DATA_ROW, worksheet.max_row + 1):
        for column in SUMMARY_FIGURE_COLUMNS:
            set_cell_font_color(worksheet.cell(row_number, column), SUMMARY_FONT_BLACK)


def rebuild_summary_sheet(
    worksheet: Any,
    start_date: date,
    end_date: date,
    metrics: dict[int, ProjectMetrics],
    metadata: dict[int, ProjectMetadata],
) -> None:
    validate_summary_layout(worksheet)
    missing_projects = sorted(set(metrics) - set(metadata))
    if missing_projects:
        raise InputValidationError(
            "Eligible ECM project code(s) are absent from the SMM readiness sheet: "
            + ", ".join(map(str, missing_projects))
        )

    # Remove inherited template colors before capturing styles for the new rows.
    normalize_summary_figure_font_colors(worksheet)

    first_styles = capture_row_styles(worksheet, SUMMARY_FIRST_DATA_ROW)
    body_row = min(SUMMARY_FIRST_DATA_ROW + 1, worksheet.max_row)
    body_styles = capture_row_styles(worksheet, body_row)
    last_styles = capture_row_styles(worksheet, worksheet.max_row)
    first_height = worksheet.row_dimensions[SUMMARY_FIRST_DATA_ROW].height
    body_height = worksheet.row_dimensions[body_row].height
    last_height = worksheet.row_dimensions[worksheet.max_row].height

    if worksheet.max_row >= SUMMARY_FIRST_DATA_ROW:
        worksheet.delete_rows(
            SUMMARY_FIRST_DATA_ROW,
            worksheet.max_row - SUMMARY_FIRST_DATA_ROW + 1,
        )

    label = period_label(start_date, end_date)
    worksheet.cell(2, 5).value = f"Number of\nMethod Statement\n  Record in SMM\n({label})"
    worksheet.cell(2, 8).value = f"Number of\nMethod Statement\nRecord (Outbound) in ECM\n({label})"

    project_codes = sorted(metrics)
    for index, project_code in enumerate(project_codes):
        row_number = SUMMARY_FIRST_DATA_ROW + index
        if len(project_codes) == 1:
            styles = first_styles
            height = first_height
        elif index == 0:
            styles = first_styles
            height = first_height
        elif index == len(project_codes) - 1:
            styles = last_styles
            height = last_height
        else:
            styles = body_styles
            height = body_height
        apply_row_styles(worksheet, row_number, styles)
        worksheet.row_dimensions[row_number].height = height
        worksheet.row_dimensions[row_number].hidden = False

        project = metadata[project_code]
        values = {
            2: project_display_value(project),
            3: project.business_unit,
            4: None,
            5: metrics[project_code].draft_via_kms,
            6: metrics[project_code].draft_not_via_kms,
            7: f"=E{row_number}+F{row_number}",
            8: metrics[project_code].outbound_ecm,
        }
        for column, value in values.items():
            cell = worksheet.cell(row_number, column)
            cell.value = value
            if column in SUMMARY_FIGURE_COLUMNS:
                set_cell_font_color(cell, SUMMARY_FONT_BLACK)

        if metrics[project_code].draft_not_via_kms > 0:
            set_cell_font_color(worksheet.cell(row_number, 6), SUMMARY_FONT_RED)

    apply_summary_column_formatting(worksheet)


def create_smm_workbook(
    ecm_record_path: Path,
    smm_template_path: Path,
    smm_record_path: Path,
    backup_path: Path,
    start_date: date,
    end_date: date,
    ecm_worksheet_name: str,
    summary_template_sheet: str,
) -> tuple[str, int]:
    if not smm_template_path.is_file():
        raise FileNotFoundError(f"SMM template workbook was not found: {smm_template_path}")
    if smm_template_path.resolve() == smm_record_path.resolve():
        raise InputValidationError("The SMM template and working workbook paths must be different.")
    if smm_record_path.resolve() == backup_path.resolve():
        raise InputValidationError("The SMM working workbook and backup paths must be different.")
    if not ecm_record_path.is_file():
        raise FileNotFoundError(f"ECM record workbook was not found: {ecm_record_path}")

    metrics = aggregate_ecm_record(ecm_record_path, ecm_worksheet_name)
    if not metrics:
        raise InputValidationError("The ECM review contains no rows with Count (Y/N) equal to Y.")

    # J/K are completed after Step 1b. Store the presentation value only after
    # those manual decisions have been validated by the aggregation pass.
    write_display_values(ecm_record_path, ecm_worksheet_name)

    target_sheet_name = summary_sheet_name(start_date, end_date)
    smm_record_path.parent.mkdir(parents=True, exist_ok=True)
    backup_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        if not smm_record_path.is_file():
            shutil.copy2(smm_template_path, smm_record_path)
    except PermissionError as error:
        raise OSError(
            f"Cannot initialize the SMM working workbook; close it if it is open: {smm_record_path}"
        ) from error

    workbook = load_workbook(smm_record_path)
    try:
        if summary_template_sheet not in workbook.sheetnames:
            raise InputValidationError(
                f"Summary layout sheet {summary_template_sheet!r} was not found in {smm_record_path.name}."
            )
        metadata = read_project_metadata(workbook)
        source_worksheet = workbook[summary_template_sheet]
        validate_summary_layout(source_worksheet)
        if target_sheet_name in workbook.sheetnames:
            workbook.remove(workbook[target_sheet_name])
        worksheet = workbook.copy_worksheet(source_worksheet)
        worksheet.title = target_sheet_name
        rebuild_summary_sheet(worksheet, start_date, end_date, metrics, metadata)

        for sheet in workbook.worksheets:
            sheet.sheet_view.tabSelected = False
        worksheet.sheet_view.tabSelected = True
        workbook.active = workbook.index(worksheet)
        try:
            workbook.save(smm_record_path)
        except PermissionError as error:
            raise OSError(
                f"Cannot save the SMM working workbook; close it if it is open: {smm_record_path}"
            ) from error
    except Exception:
        workbook.close()
        raise
    else:
        workbook.close()
    try:
        shutil.copy2(smm_record_path, backup_path)
    except PermissionError as error:
        raise OSError(
            f"Cannot create the SMM backup; close it if it is open: {backup_path}"
        ) from error
    return target_sheet_name, len(metrics)


def print_metrics_audit(metrics: dict[int, ProjectMetrics]) -> None:
    """Print per-project values needed to audit Mattex deductions in run logs."""
    for project_code in sorted(metrics):
        project = metrics[project_code]
        print(
            f"Project {project_code}: outbound raw={project.raw_outbound_ecm}; "
            f"Mattex deduction={project.mattex_added}; outbound final={project.outbound_ecm}"
        )


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Step 2: update the SMM working workbook with a period Summary sheet."
    )
    parser.add_argument("--start-date", required=True, type=parse_date)
    parser.add_argument("--end-date", required=True, type=parse_date)
    parser.add_argument(
        "--ecm-record",
        type=Path,
        default=DEFAULT_ECM_RECORD,
        help="Manually reviewed ECM workbook. Defaults to output/ECM_Record.xlsx.",
    )
    parser.add_argument(
        "--smm-template",
        type=Path,
        default=DEFAULT_SMM_TEMPLATE,
        help="Immutable SMM template used only to initialize the working workbook.",
    )
    parser.add_argument(
        "--smm-record",
        type=Path,
        default=DEFAULT_SMM_RECORD,
        help="Persistent SMM working workbook to update.",
    )
    parser.add_argument("--ecm-sheet", help="ECM review sheet. Defaults to the reporting-period name.")
    parser.add_argument(
        "--summary-template-sheet",
        default=DEFAULT_SUMMARY_TEMPLATE_SHEET,
        help="Existing SMM period sheet to copy as the Summary layout.",
    )
    parser.add_argument(
        "--backup",
        type=Path,
        help="Date-stamped backup path. Defaults to a reporting-period backup in output.",
    )
    args = parser.parse_args()
    if args.end_date < args.start_date:
        parser.error("--end-date cannot be earlier than --start-date.")
    return args


def main() -> int:
    args = parse_arguments()
    backup_path = args.backup or (
        DEFAULT_OUTPUT_DIRECTORY
        / f"SMM_KMS_PM_Master_Document_{args.start_date:%Y-%m-%d}_to_{args.end_date:%Y-%m-%d}.xlsx"
    )
    ecm_worksheet_name = args.ecm_sheet or ecm_sheet_name(args.start_date, args.end_date)
    metrics = aggregate_ecm_record(args.ecm_record, ecm_worksheet_name)
    target_sheet_name, project_count = create_smm_workbook(
        ecm_record_path=args.ecm_record,
        smm_template_path=args.smm_template,
        smm_record_path=args.smm_record,
        backup_path=backup_path,
        start_date=args.start_date,
        end_date=args.end_date,
        ecm_worksheet_name=ecm_worksheet_name,
        summary_template_sheet=args.summary_template_sheet,
    )

    print(f"ECM record: {args.ecm_record.resolve()}")
    print(f"ECM sheet: {ecm_worksheet_name}")
    print(f"SMM template: {args.smm_template.resolve()}")
    print(f"Summary layout: {args.summary_template_sheet}")
    print(f"New Summary sheet: {target_sheet_name}")
    print_metrics_audit(metrics)
    print(f"SMM working workbook: {args.smm_record.resolve()}")
    print(f"Backup: {backup_path.resolve()}")
    print(f"Summary projects: {project_count}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (FileNotFoundError, InputValidationError, OSError, ValueError) as error:
        print(f"Error: {error}", file=sys.stderr)
        raise SystemExit(1)
