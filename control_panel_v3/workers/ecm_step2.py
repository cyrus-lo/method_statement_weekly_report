"""Create a V3 SMM period summary from a manually reviewed ECM worksheet.

Both workbook updates are staged and validated before publication. If any of the
three final replacements (ECM, SMM, period backup) fails, previous files are
restored or the recovery copies are retained for operator recovery.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import sys
import tempfile
from collections import defaultdict
from copy import copy
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

CONTROL_PANEL_ROOT = Path(__file__).resolve().parents[1]
READINESS_SHEET_NAME = "Project Realiness Chris-In Use"
SUMMARY_HEADER_ROWS = 4
SUMMARY_FIRST_DATA_ROW = 5
SUMMARY_FIGURE_COLUMNS = (5, 6, 7, 8)
SUMMARY_COLUMN_WIDTHS = {
    "B": 93.1875,
    "C": 16.46375,
    "E": 20.46375,
    "F": 20.46375,
    "G": 20.46375,
    "H": 27.46375,
}
SUMMARY_FONT_BLACK = "FF000000"
SUMMARY_FONT_RED = "FFFF0000"
ECM_HEADERS = {
    1: "Project Code",
    8: "SMM ID",
    9: "Count (Y/N)",
    10: "Use SMM (Y/N)",
    11: "Use KMS (Y/N)",
    13: "Remark",
}
PROJECT_CODE_PATTERN = re.compile(r"^\s*(\d+)\b")
MATTEX_ADDED_REMARKS = ("added by mattex", "add by chris's team")
CANONICAL_PROJECT_DISPLAY = {1016: "1016 - 1/WSD/19 (1007)"}


class Step2Error(ValueError):
    """Invalid Step 2 workbook data or V3 path contract."""


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


def normalise_text(value: Any) -> str:
    return str(value or "").strip()


def is_mattex_added(value: Any) -> bool:
    remark = " ".join(normalise_text(value).casefold().replace("’", "'").replace("‘", "'").split())
    return any(marker in remark for marker in MATTEX_ADDED_REMARKS)


def normalise_project_code(value: Any) -> int | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        code = value
    elif isinstance(value, float) and value.is_integer():
        code = int(value)
    else:
        match = PROJECT_CODE_PATTERN.match(normalise_text(value))
        if not match:
            return None
        code = int(match.group(1))
    return 1016 if code == 1007 else code


def project_display_value(project: ProjectMetadata) -> Any:
    return CANONICAL_PROJECT_DISPLAY.get(project.project_code, project.project)


def normalise_flag(value: Any, *, field: str, row_number: int, allow_blank: bool) -> str:
    text = normalise_text(value).upper()
    allowed = {"Y", "N"} | ({""} if allow_blank else set())
    if text not in allowed:
        expected = "Y, N, or blank" if allow_blank else "Y or N"
        raise Step2Error(f"ECM row {row_number}: {field} must be {expected}; got {value!r}.")
    return text


def period_label(start_date: date, end_date: date) -> str:
    if start_date.year == end_date.year:
        return f"{start_date.day} {start_date:%b} - {end_date.day} {end_date:%b} {end_date.year}"
    return f"{start_date.day} {start_date:%b %Y} - {end_date.day} {end_date:%b %Y}"


def summary_sheet_name(start_date: date, end_date: date) -> str:
    name = f"(WIP) {start_date.day} {start_date:%b} - {end_date.day} {end_date:%b} {end_date.year}"
    if len(name) > 31:
        raise Step2Error(f"Generated summary worksheet name exceeds Excel's 31-character limit: {name!r}.")
    return name


def validate_ecm_headers(worksheet: Any) -> None:
    mismatches = []
    for column, expected in ECM_HEADERS.items():
        actual = normalise_text(worksheet.cell(1, column).value)
        if actual != expected:
            mismatches.append(f"{worksheet.cell(1, column).coordinate}: expected {expected!r}, got {actual!r}")
    display_heading = normalise_text(worksheet.cell(1, 14).value)
    if display_heading not in {"", "Display"}:
        mismatches.append(f"N1: expected 'Display' or blank, got {display_heading!r}")
    if mismatches:
        raise Step2Error("ECM review headers do not match: " + "; ".join(mismatches))


def _review_data_rows(worksheet: Any) -> list[int]:
    relevant_columns = (*range(1, 12), 13)
    return [
        row_number
        for row_number in range(2, worksheet.max_row + 1)
        if any(worksheet.cell(row_number, column).value not in (None, "") for column in relevant_columns)
    ]


def display_value(count: Any, use_smm: Any, use_kms: Any) -> str:
    flags = tuple(normalise_text(value).upper() for value in (count, use_smm, use_kms))
    if flags == ("Y", "Y", "Y"):
        return "via KMS"
    if flags == ("Y", "Y", "N"):
        return "w/o KMS"
    return ""


def _has_review_data(worksheet: Any, row_number: int) -> bool:
    return any(worksheet.cell(row_number, column).value is not None for column in range(1, 9))


def aggregate_ecm_record(ecm_record_path: Path, worksheet_name: str) -> tuple[dict[int, ProjectMetrics], dict[int, str]]:
    if not ecm_record_path.is_file():
        raise FileNotFoundError(f"ECM review workbook was not found: {ecm_record_path}")
    workbook = load_workbook(ecm_record_path, read_only=True, data_only=True)
    try:
        if worksheet_name not in workbook.sheetnames:
            raise Step2Error(f"ECM worksheet {worksheet_name!r} was not found in {ecm_record_path.name}.")
        worksheet = workbook[worksheet_name]
        validate_ecm_headers(worksheet)
        metrics: dict[int, ProjectMetrics] = defaultdict(ProjectMetrics)
        display_values: dict[int, str] = {}
        for row_number in _review_data_rows(worksheet):
            count_flag = normalise_flag(
                worksheet.cell(row_number, 9).value,
                field="Count (Y/N)",
                row_number=row_number,
                allow_blank=False,
            )
            if _has_review_data(worksheet, row_number):
                display_values[row_number] = display_value(
                    worksheet.cell(row_number, 9).value,
                    worksheet.cell(row_number, 10).value,
                    worksheet.cell(row_number, 11).value,
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
                raise Step2Error(f"ECM row {row_number}: Project Code is required when Count (Y/N) is Y.")
            mattex_added = is_mattex_added(worksheet.cell(row_number, 13).value)
            project = metrics[project_code]
            if use_smm == "Y" and use_kms == "Y":
                project.draft_via_kms += 1
            elif use_smm == "Y" and use_kms == "N":
                project.draft_not_via_kms += 1
            else:
                if use_kms == "Y":
                    raise Step2Error(f"ECM row {row_number}: Use KMS cannot be Y when Use SMM is not Y.")
                if mattex_added:
                    raise Step2Error(f"ECM row {row_number}: a Mattex-added Remark cannot be used when Use SMM is N.")
                project.not_using_smm += 1
                continue
            if mattex_added:
                project.mattex_added += 1
        if not metrics:
            raise Step2Error("The ECM review contains no rows with Count (Y/N) equal to Y.")
        return dict(metrics), display_values
    finally:
        workbook.close()


def read_project_metadata(workbook: Any) -> dict[int, ProjectMetadata]:
    if READINESS_SHEET_NAME not in workbook.sheetnames:
        raise Step2Error(f"Readiness sheet {READINESS_SHEET_NAME!r} was not found in the V3 readiness workbook.")
    worksheet = workbook[READINESS_SHEET_NAME]
    metadata: dict[int, ProjectMetadata] = {}
    for row_number in range(2, worksheet.max_row + 1):
        project_value = worksheet.cell(row_number, 3).value
        project_code = normalise_project_code(project_value)
        if project_code is None:
            continue
        if project_code in metadata:
            raise Step2Error(f"Duplicate project code {project_code} in V3 readiness sheet row {row_number}.")
        metadata[project_code] = ProjectMetadata(
            project_code=project_code,
            project=project_value,
            business_unit=worksheet.cell(row_number, 4).value,
        )
    if not metadata:
        raise Step2Error("No project codes were found in the V3 readiness sheet.")
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
            mismatches.append(f"{worksheet.cell(row, column).coordinate}: expected {expected_heading!r}, got {actual!r}")
    outbound_header = " ".join(normalise_text(worksheet.cell(2, 8).value).split())
    if "Record (Outbound) in ECM" not in outbound_header:
        mismatches.append(f"H2: expected an outbound ECM heading, got {outbound_header!r}")
    if mismatches:
        raise Step2Error("Summary layout does not match: " + "; ".join(mismatches))


def capture_row_styles(worksheet: Any, row_number: int) -> list[Any]:
    return [copy(worksheet.cell(row_number, column)._style) for column in range(1, 9)]


def apply_row_styles(worksheet: Any, row_number: int, styles: list[Any]) -> None:
    for column, style in enumerate(styles, start=1):
        worksheet.cell(row_number, column)._style = copy(style)


def set_cell_font_color(cell: Any, color: str) -> None:
    font = copy(cell.font)
    font.color = color
    cell.font = font


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
        raise Step2Error(
            "Eligible ECM project code(s) are absent from the V3 readiness sheet: "
            + ", ".join(map(str, missing_projects))
        )

    for row_number in range(SUMMARY_FIRST_DATA_ROW, worksheet.max_row + 1):
        for column in SUMMARY_FIGURE_COLUMNS:
            set_cell_font_color(worksheet.cell(row_number, column), SUMMARY_FONT_BLACK)

    first_styles = capture_row_styles(worksheet, SUMMARY_FIRST_DATA_ROW)
    body_row = min(SUMMARY_FIRST_DATA_ROW + 1, worksheet.max_row)
    body_styles = capture_row_styles(worksheet, body_row)
    last_styles = capture_row_styles(worksheet, worksheet.max_row)
    first_height = worksheet.row_dimensions[SUMMARY_FIRST_DATA_ROW].height
    body_height = worksheet.row_dimensions[body_row].height
    last_height = worksheet.row_dimensions[worksheet.max_row].height
    if worksheet.max_row >= SUMMARY_FIRST_DATA_ROW:
        worksheet.delete_rows(SUMMARY_FIRST_DATA_ROW, worksheet.max_row - SUMMARY_FIRST_DATA_ROW + 1)

    label = period_label(start_date, end_date)
    worksheet.cell(2, 5).value = f"Number of\nMethod Statement\n  Record in SMM\n({label})"
    worksheet.cell(2, 8).value = f"Number of\nMethod Statement\nRecord (Outbound) in ECM\n({label})"
    project_codes = sorted(metrics)
    for index, project_code in enumerate(project_codes):
        row_number = SUMMARY_FIRST_DATA_ROW + index
        if len(project_codes) == 1 or index == 0:
            styles, height = first_styles, first_height
        elif index == len(project_codes) - 1:
            styles, height = last_styles, last_height
        else:
            styles, height = body_styles, body_height
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

    for column, width in SUMMARY_COLUMN_WIDTHS.items():
        worksheet.column_dimensions[column].width = width
    worksheet.column_dimensions["D"].hidden = True
    worksheet.sheet_view.zoomScale = 85


def _ensure_v3_path(label: str, path: Path) -> Path:
    resolved = path.resolve()
    try:
        resolved.relative_to(CONTROL_PANEL_ROOT.resolve())
    except ValueError as error:
        raise Step2Error(f"{label} must remain under the V3 control-panel root: {resolved}") from error
    return resolved


def _temporary_path(directory: Path, prefix: str) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(prefix=prefix, suffix=".xlsx", dir=directory, delete=False) as stream:
        return Path(stream.name)


def _file_digest(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _write_display_values(ecm_path: Path, staged_path: Path, worksheet_name: str, display_values: dict[int, str]) -> None:
    workbook = load_workbook(ecm_path)
    try:
        if worksheet_name not in workbook.sheetnames:
            raise Step2Error(f"ECM worksheet {worksheet_name!r} was not found in {ecm_path.name}.")
        worksheet = workbook[worksheet_name]
        validate_ecm_headers(worksheet)
        if not normalise_text(worksheet.cell(1, 14).value):
            worksheet.cell(1, 14).value = "Display"
        for row_number, value in display_values.items():
            worksheet.cell(row_number, 14).value = value or None
        workbook.save(staged_path)
    finally:
        workbook.close()


def _verify_staged_workbooks(ecm_path: Path, ecm_sheet: str, smm_path: Path, summary_sheet: str, backup_path: Path, display_values: dict[int, str], project_count: int) -> None:
    ecm_workbook = load_workbook(ecm_path, read_only=True, data_only=True)
    try:
        if ecm_sheet not in ecm_workbook.sheetnames:
            raise Step2Error("Staged ECM workbook is missing the selected period worksheet.")
        worksheet = ecm_workbook[ecm_sheet]
        validate_ecm_headers(worksheet)
        for row_number, expected in display_values.items():
            actual = worksheet.cell(row_number, 14).value
            if actual not in ({expected, None, ""} if not expected else {expected}):
                raise Step2Error(f"Staged ECM Display value is incorrect at N{row_number}.")
    finally:
        ecm_workbook.close()
    for path in (smm_path, backup_path):
        workbook = load_workbook(path, read_only=False, data_only=False)
        try:
            if summary_sheet not in workbook.sheetnames:
                raise Step2Error(f"Staged SMM workbook is missing the generated sheet {summary_sheet!r}.")
            worksheet = workbook[summary_sheet]
            for row_number in range(SUMMARY_FIRST_DATA_ROW, SUMMARY_FIRST_DATA_ROW + project_count):
                if worksheet.cell(row_number, 7).value != f"=E{row_number}+F{row_number}":
                    raise Step2Error(f"Staged SMM workbook is missing the expected Total formula at G{row_number}.")
            if worksheet.sheet_view.zoomScale != 85:
                raise Step2Error("Staged SMM workbook does not have the required 85% summary zoom.")
        finally:
            workbook.close()


def _verify_active_and_backup(smm_path: Path, backup_path: Path, summary_sheet: str) -> None:
    for path in (smm_path, backup_path):
        workbook = load_workbook(path, read_only=False, data_only=False)
        try:
            if workbook.active.title != summary_sheet:
                raise Step2Error(f"Workbook {path.name} does not select the generated period worksheet.")
            if not workbook[summary_sheet].sheet_view.tabSelected:
                raise Step2Error(f"Workbook {path.name} does not select the generated period worksheet tab.")
        finally:
            workbook.close()


def _restore_file(path: Path, rollback_path: Path | None, existed: bool) -> None:
    if existed:
        if rollback_path is None or not rollback_path.is_file():
            raise OSError(f"Recovery copy for {path} is missing.")
        os.replace(rollback_path, path)
    else:
        path.unlink(missing_ok=True)


def run_step2(
    ecm_record_path: Path,
    ecm_worksheet_name: str,
    smm_record_path: Path,
    smm_template_path: Path,
    readiness_workbook_path: Path,
    backup_path: Path,
    start_date: date,
    end_date: date,
    summary_template_sheet: str,
) -> tuple[str, int]:
    if end_date < start_date:
        raise Step2Error("End date cannot be earlier than start date.")
    ecm_record_path = _ensure_v3_path("ECM review workbook", ecm_record_path)
    smm_record_path = _ensure_v3_path("V3 SMM working workbook", smm_record_path)
    smm_template_path = _ensure_v3_path("V3 SMM template", smm_template_path)
    readiness_workbook_path = _ensure_v3_path("V3 readiness workbook", readiness_workbook_path)
    backup_path = _ensure_v3_path("V3 SMM period backup", backup_path)
    if len({ecm_record_path, smm_record_path, smm_template_path, readiness_workbook_path, backup_path}) != 5:
        raise Step2Error("ECM workbook, SMM workbook, SMM template, readiness workbook, and period backup paths must all be different.")
    if not ecm_record_path.is_file():
        raise FileNotFoundError(f"V3 ECM review workbook was not found: {ecm_record_path}")
    if not smm_template_path.is_file():
        raise FileNotFoundError(f"V3 SMM template was not found: {smm_template_path}")
    if not readiness_workbook_path.is_file():
        raise FileNotFoundError(f"V3 readiness workbook was not found: {readiness_workbook_path}")
    if ecm_record_path.with_name(f"~${ecm_record_path.name}").exists():
        raise Step2Error(f"Close the V3 ECM workbook before running Step 2: {ecm_record_path.name}")
    if smm_record_path.with_name(f"~${smm_record_path.name}").exists():
        raise Step2Error(f"Close the V3 SMM workbook before running Step 2: {smm_record_path.name}")
    if backup_path.with_name(f"~${backup_path.name}").exists():
        raise Step2Error(f"Close the V3 SMM period backup before running Step 2: {backup_path.name}")

    target_sheet_name = summary_sheet_name(start_date, end_date)
    if target_sheet_name.casefold() == summary_template_sheet.casefold():
        raise Step2Error("Generated period sheet cannot replace the configured summary layout source.")
    metrics, display_values = aggregate_ecm_record(ecm_record_path, ecm_worksheet_name)
    ecm_digest_before = _file_digest(ecm_record_path)
    smm_digest_before = _file_digest(smm_record_path) if smm_record_path.is_file() else None
    template_digest_before = _file_digest(smm_template_path)
    readiness_digest_before = _file_digest(readiness_workbook_path)
    backup_digest_before = _file_digest(backup_path) if backup_path.is_file() else None
    source_workbook = load_workbook(smm_template_path)
    try:
        if summary_template_sheet not in source_workbook.sheetnames:
            raise Step2Error(f"Summary layout sheet {summary_template_sheet!r} was not found in the V3 SMM template {smm_template_path.name}.")
        validate_summary_layout(source_workbook[summary_template_sheet])
    finally:
        source_workbook.close()
    readiness_workbook = load_workbook(readiness_workbook_path, read_only=True, data_only=True)
    try:
        if READINESS_SHEET_NAME not in readiness_workbook.sheetnames:
            raise Step2Error(f"Readiness sheet {READINESS_SHEET_NAME!r} was not found in the V3 readiness workbook.")
        metadata = read_project_metadata(readiness_workbook)
    finally:
        readiness_workbook.close()
    missing_projects = sorted(set(metrics) - set(metadata))
    if missing_projects:
        raise Step2Error(
            "Eligible ECM project code(s) are absent from the V3 readiness sheet: "
            + ", ".join(map(str, missing_projects))
        )

    ecm_staged_path: Path | None = None
    smm_staged_path: Path | None = None
    backup_staged_path: Path | None = None
    ecm_rollback_path: Path | None = None
    smm_rollback_path: Path | None = None
    backup_rollback_path: Path | None = None
    preserve_recovery = False
    ecm_existed = ecm_record_path.exists()
    smm_existed = smm_record_path.exists()
    backup_existed = backup_path.exists()
    try:
        ecm_staged_path = _temporary_path(ecm_record_path.parent, ".v3-step2-ecm-")
        _write_display_values(ecm_record_path, ecm_staged_path, ecm_worksheet_name, display_values)

        smm_staged_path = _temporary_path(smm_record_path.parent, ".v3-step2-smm-")
        shutil.copy2(smm_record_path if smm_existed else smm_template_path, smm_staged_path)
        smm_workbook = load_workbook(smm_staged_path)
        try:
            if summary_template_sheet not in smm_workbook.sheetnames:
                raise Step2Error(f"Summary layout sheet {summary_template_sheet!r} was not found in the staged V3 SMM workbook.")
            template_workbook = load_workbook(smm_template_path)
            try:
                source_sheet = template_workbook[summary_template_sheet]
                validate_summary_layout(source_sheet)
                template_sheet = source_sheet
                if target_sheet_name.casefold() == summary_template_sheet.casefold():
                    raise Step2Error("Generated period sheet cannot replace the configured summary layout source.")
                existing_target = next((name for name in smm_workbook.sheetnames if name.casefold() == target_sheet_name.casefold()), None)
                if existing_target is not None:
                    if existing_target.casefold() == summary_template_sheet.casefold() or not existing_target.casefold().startswith("(wip) "):
                        raise Step2Error("Generated period sheet cannot replace the configured summary layout source or an unrelated worksheet.")
                    del smm_workbook[existing_target]
                new_sheet = smm_workbook.create_sheet(target_sheet_name)
                for row in template_sheet.iter_rows():
                    for source_cell in row:
                        target_cell = new_sheet.cell(source_cell.row, source_cell.column, source_cell.value)
                        if source_cell.has_style:
                            target_cell._style = copy(source_cell._style)
                        if source_cell.number_format:
                            target_cell.number_format = source_cell.number_format
                        if source_cell.alignment:
                            target_cell.alignment = copy(source_cell.alignment)
                        if source_cell.protection:
                            target_cell.protection = copy(source_cell.protection)
                for row_number, dimension in template_sheet.row_dimensions.items():
                    new_sheet.row_dimensions[row_number].height = dimension.height
                    new_sheet.row_dimensions[row_number].hidden = dimension.hidden
                for column_letter, dimension in template_sheet.column_dimensions.items():
                    new_sheet.column_dimensions[column_letter].width = dimension.width
                    new_sheet.column_dimensions[column_letter].hidden = dimension.hidden
                    new_sheet.column_dimensions[column_letter].bestFit = dimension.bestFit
                for merged_range in template_sheet.merged_cells.ranges:
                    new_sheet.merge_cells(str(merged_range))
                new_sheet.freeze_panes = template_sheet.freeze_panes
                new_sheet.sheet_view.showGridLines = template_sheet.sheet_view.showGridLines
                new_sheet.sheet_view.zoomScale = template_sheet.sheet_view.zoomScale
                new_sheet.sheet_properties.pageSetUpPr.fitToPage = template_sheet.sheet_properties.pageSetUpPr.fitToPage
                new_sheet.page_margins = copy(template_sheet.page_margins)
                new_sheet.page_setup = copy(template_sheet.page_setup)
                new_sheet.print_options = copy(template_sheet.print_options)
            finally:
                template_workbook.close()
            rebuild_summary_sheet(new_sheet, start_date, end_date, metrics, metadata)
            for sheet in smm_workbook.worksheets:
                sheet.sheet_view.tabSelected = False
            new_sheet.sheet_view.tabSelected = True
            smm_workbook.active = smm_workbook.index(new_sheet)
            smm_workbook.save(smm_staged_path)
        finally:
            smm_workbook.close()

        backup_path.parent.mkdir(parents=True, exist_ok=True)
        backup_staged_path = _temporary_path(backup_path.parent, ".v3-step2-backup-")
        shutil.copy2(smm_staged_path, backup_staged_path)
        _verify_staged_workbooks(ecm_staged_path, ecm_worksheet_name, smm_staged_path, target_sheet_name, backup_staged_path, display_values, len(metrics))
        if _file_digest(ecm_record_path) != ecm_digest_before:
            raise Step2Error("V3 ECM workbook changed during Step 2 preparation; refusing to commit a stale staged workbook.")
        if smm_record_path.is_file() != smm_existed or (smm_existed and _file_digest(smm_record_path) != smm_digest_before):
            raise Step2Error("V3 SMM workbook changed during Step 2 preparation; refusing to commit a stale staged workbook.")
        if _file_digest(smm_template_path) != template_digest_before:
            raise Step2Error("V3 SMM source template changed during Step 2 preparation; refusing to commit.")
        if _file_digest(readiness_workbook_path) != readiness_digest_before:
            raise Step2Error("V3 readiness workbook changed during Step 2 preparation; refusing to commit stale readiness data.")
        if backup_path.is_file() != backup_existed or (backup_existed and _file_digest(backup_path) != backup_digest_before):
            raise Step2Error("V3 SMM period backup changed during Step 2 preparation; refusing to overwrite it.")

        ecm_rollback_path = _temporary_path(ecm_record_path.parent, ".v3-step2-recovery-ecm-")
        shutil.copy2(ecm_record_path, ecm_rollback_path)
        if smm_existed:
            smm_rollback_path = _temporary_path(smm_record_path.parent, ".v3-step2-recovery-smm-")
            shutil.copy2(smm_record_path, smm_rollback_path)
        if backup_existed:
            backup_rollback_path = _temporary_path(backup_path.parent, ".v3-step2-recovery-backup-")
            shutil.copy2(backup_path, backup_rollback_path)

        try:
            if _file_digest(readiness_workbook_path) != readiness_digest_before:
                raise Step2Error("V3 readiness workbook changed before Step 2 commit; refusing to publish results based on stale readiness data.")
            os.replace(ecm_staged_path, ecm_record_path)
            ecm_staged_path = None
            os.replace(smm_staged_path, smm_record_path)
            smm_staged_path = None
            os.replace(backup_staged_path, backup_path)
            backup_staged_path = None
            _verify_active_and_backup(smm_record_path, backup_path, target_sheet_name)
        except Exception as commit_error:
            rollback_errors: list[str] = []
            for path, rollback, existed in (
                (ecm_record_path, ecm_rollback_path, ecm_existed),
                (smm_record_path, smm_rollback_path, smm_existed),
                (backup_path, backup_rollback_path, backup_existed),
            ):
                try:
                    _restore_file(path, rollback, existed)
                    if path == ecm_record_path:
                        ecm_rollback_path = None
                    elif path == smm_record_path:
                        smm_rollback_path = None
                    else:
                        backup_rollback_path = None
                except Exception as rollback_error:
                    rollback_errors.append(f"{path}: {rollback_error}")
            if rollback_errors:
                preserve_recovery = True
                raise Step2Error(
                    f"Step 2 commit failed ({commit_error}); rollback also failed. Recovery evidence retained: "
                    + "; ".join(rollback_errors)
                ) from commit_error
            raise

        try:
            for path in (ecm_rollback_path, smm_rollback_path, backup_rollback_path):
                if path is not None:
                    path.unlink(missing_ok=True)
        except OSError as cleanup_error:
            preserve_recovery = True
            raise Step2Error(f"Step 2 committed, but rollback evidence could not be cleaned up; preserve recovery files and verify outputs: {cleanup_error}") from cleanup_error
        ecm_rollback_path = smm_rollback_path = backup_rollback_path = None
        print(json.dumps({"event": "step2_complete", "projects": len(metrics), "worksheet": target_sheet_name}, sort_keys=True))
        print(f"Updated ECM workbook: {ecm_record_path}")
        print(f"Updated V3 SMM workbook: {smm_record_path}")
        print(f"Post-Step-2 SMM backup: {backup_path}")
        for project_code in sorted(metrics):
            project = metrics[project_code]
            print(
                f"Project {project_code}: outbound raw={project.raw_outbound_ecm}; "
                f"Mattex deduction={project.mattex_added}; outbound final={project.outbound_ecm}"
            )
        return target_sheet_name, len(metrics)
    finally:
        for path in (ecm_staged_path, smm_staged_path, backup_staged_path):
            if path is not None:
                path.unlink(missing_ok=True)
        if not preserve_recovery:
            for path in (ecm_rollback_path, smm_rollback_path, backup_rollback_path):
                if path is not None:
                    path.unlink(missing_ok=True)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Create a V3 SMM summary from a reviewed ECM worksheet.")
    parser.add_argument("--start-date", required=True, type=date.fromisoformat)
    parser.add_argument("--end-date", required=True, type=date.fromisoformat)
    parser.add_argument("--ecm-workbook", required=True, type=Path)
    parser.add_argument("--ecm-sheet", required=True)
    parser.add_argument("--smm-workbook", required=True, type=Path)
    parser.add_argument("--smm-template", required=True, type=Path)
    parser.add_argument("--readiness-workbook", required=True, type=Path)
    parser.add_argument("--summary-layout-sheet", required=True)
    parser.add_argument("--backup", required=True, type=Path)
    args = parser.parse_args(argv)
    run_step2(
        args.ecm_workbook,
        args.ecm_sheet,
        args.smm_workbook,
        args.smm_template,
        args.readiness_workbook,
        args.backup,
        args.start_date,
        args.end_date,
        args.summary_layout_sheet,
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (FileNotFoundError, OSError, Step2Error, ValueError) as error:
        print(f"V3 Step 2 failed: {error}", file=sys.stderr)
        raise SystemExit(1)
