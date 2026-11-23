"""Create the V3 ECM review worksheet from one exact-period CSV handoff."""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import shutil
import sys
import tempfile
from copy import copy
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import Border

from .properties_csv import CsvValidationError, validate_properties_csv
from .reference_normalization import outgoing_reference_without_revision

CONTROL_PANEL_ROOT = Path(__file__).resolve().parents[1]

REVIEW_HEADERS = (
    "Project Code", "Document title", "Document date", "Name", "Document revision",
    "Modified By", "Major Version", "SMM ID", "Count (Y/N)", "Use SMM (Y/N)", "Use KMS (Y/N)",
)
READINESS_SHEET_NAME = "Project Realiness Chris-In Use"
TEMPLATE_SHEETS = ("ECM_template", "ECM_1-12 Jul (WIP)", "ECM_1-12 Jul (OK)")
OUTGOING_REFERENCE_HEADER = "Outgoing ref. no. (without revision)"
OUTGOING_REFERENCE_COLUMN = 16


class Step1aError(ValueError):
    """Invalid Step 1a inputs or workbook contract."""


@dataclass
class EcmRecord:
    source_row: int
    document_title: str
    document_date: datetime
    name: str
    document_revision: str
    modified_by: str
    major_version: int
    smm_id: str
    modified_on: datetime | None
    outgoing_reference: str
    project_code: int | None = None
    reasons: list[str] = field(default_factory=list)
    count: str = ""

    @property
    def normalized_name(self) -> str:
        return self.name.strip().casefold()


def period_sheet_name(start_date: date, end_date: date) -> str:
    if start_date.month == end_date.month:
        return f"ECM_{start_date.day}-{end_date.day} {end_date:%b} (WIP)"
    return f"ECM_{start_date.day} {start_date:%b}-{end_date.day} {end_date:%b} (WIP)"


def _parse_datetime(value: str, name: str, row_number: int, *, optional: bool = False) -> datetime | None:
    text = value.strip()
    if optional and not text:
        return None
    for fmt in ("%m/%d/%Y, %I:%M %p", "%m/%d/%Y %I:%M %p"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    raise Step1aError(f"CSV row {row_number}: {name} must use M/D/YYYY, H:MM AM/PM; got {text!r}.")


def _project_code(name: str, outgoing_reference: str) -> int | None:
    pattern = re.compile(r"^[^/]*/\s*(\d+)\s*/")
    for value in (name, outgoing_reference):
        match = pattern.match(value.strip())
        if match:
            return int(match.group(1))
    name_parts = re.findall(r"[A-Za-z0-9]+", re.sub(r"\.(pdf|docx?|xlsx?|xlsm)$", "", name.strip(), flags=re.I))
    ref_parts = re.findall(r"[A-Za-z0-9]+", re.sub(r"\.(pdf|docx?|xlsx?|xlsm)$", "", outgoing_reference.strip(), flags=re.I))
    if len(name_parts) >= 2 and name_parts == ref_parts and name_parts[1].isdigit():
        return int(name_parts[1])
    return None


def _canonical_project_code(project_code: int | None) -> int | None:
    return 1016 if project_code == 1007 else project_code


def _read_records(csv_path: Path, start: date, end: date) -> list[EcmRecord]:
    if validate_properties_csv(csv_path) < 0:
        raise Step1aError("CSV validation failed.")
    records: list[EcmRecord] = []
    with csv_path.open("r", encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream)
        for source_row, raw in enumerate(reader, start=2):
            row = {str(key or "").strip(): str(value or "").strip() for key, value in raw.items()}
            document_date = _parse_datetime(row["Document date"], "Document date", source_row)
            assert document_date is not None
            if not start <= document_date.date() <= end:
                continue
            try:
                major_version = int(row["Major Version"])
            except ValueError as error:
                raise Step1aError(f"CSV row {source_row}: Major Version must be an integer.") from error
            modified_on = _parse_datetime(row["Modified On"], "Modified On", source_row, optional=True)
            records.append(EcmRecord(
                source_row, row["Document title"], document_date, row["Name"], row["Document revision"],
                row["Modified By"], major_version, row["SMM ID"], modified_on, row["Outgoing ref. no."],
                _canonical_project_code(_project_code(row["Name"], row["Outgoing ref. no."])),
            ))
    if not records:
        raise Step1aError("The exact-period CSV contains no records within the selected date range.")
    return records


def _read_readiness(path: Path) -> dict[int, bool]:
    if not path.is_file():
        raise FileNotFoundError(f"V3 readiness workbook was not found: {path}")
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if READINESS_SHEET_NAME not in workbook.sheetnames:
            raise Step1aError(f"Readiness sheet {READINESS_SHEET_NAME!r} is missing.")
        sheet = workbook[READINESS_SHEET_NAME]
        readiness: dict[int, bool] = {}
        for row_number in range(2, sheet.max_row + 1):
            raw_code = sheet.cell(row_number, 3).value
            match = re.match(r"^\s*(\d+)\b", str(raw_code or ""))
            if not match:
                continue
            code = int(match.group(1))
            if code in readiness:
                raise Step1aError(f"Duplicate project code {code} in readiness row {row_number}.")
            readiness[code] = str(sheet.cell(row_number, 5).value or "").strip().casefold() not in {"no", "n"}
        if not readiness:
            raise Step1aError("No project codes were found in readiness column C.")
        return readiness
    finally:
        workbook.close()


def apply_preliminary_rules(records: list[EcmRecord], readiness: dict[int, bool]) -> None:
    highest_versions: dict[str, int] = {}
    for record in records:
        record.reasons.clear()
        highest_versions[record.normalized_name] = max(highest_versions.get(record.normalized_name, record.major_version), record.major_version)
    candidates: dict[str, list[EcmRecord]] = {}
    for record in records:
        if record.project_code is None:
            record.reasons.append("unresolved_project_code")
        else:
            canonical = 1016 if record.project_code == 1007 else record.project_code
            if not readiness.get(canonical, False):
                record.reasons.append("project_not_ready")
        if record.major_version < highest_versions[record.normalized_name]:
            record.reasons.append("superseded_major_version")
        else:
            candidates.setdefault(record.normalized_name, []).append(record)
    for group in candidates.values():
        if len(group) > 1:
            latest = max(group, key=lambda item: (item.modified_on or datetime.min, -item.source_row))
            for item in group:
                if item is not latest:
                    item.reasons.append("duplicate_max_major_version")
    for record in records:
        record.count = "N" if record.reasons else ""


def _copy_row_style(sheet: Any, source_row: int, target_row: int) -> None:
    for column in range(1, sheet.max_column + 1):
        source, target = sheet.cell(source_row, column), sheet.cell(target_row, column)
        if source.has_style:
            target._style = copy(source._style)
        target.number_format = source.number_format
        target.alignment = copy(source.alignment)
        target.fill = copy(source.fill)


def _build_updated_workbook(source_path: Path, start: date, end: date, records: list[EcmRecord], temporary_path: Path) -> None:
    workbook = load_workbook(source_path)
    try:
        template = next((workbook[name] for name in TEMPLATE_SHEETS if name in workbook.sheetnames), None)
        if template is None:
            raise Step1aError(f"Review worksheet template must be one of {TEMPLATE_SHEETS!r}.")
        headers = tuple(str(template.cell(1, col).value or "").strip() for col in range(1, 12))
        if headers != REVIEW_HEADERS:
            raise Step1aError(f"Review worksheet headers A:K do not match the required layout: {headers!r}.")
        sheet_name = period_sheet_name(start, end)
        if sheet_name in workbook.sheetnames:
            del workbook[sheet_name]
        sheet = workbook.copy_worksheet(template)
        sheet.title = sheet_name
        sheet.cell(1, OUTGOING_REFERENCE_COLUMN).value = OUTGOING_REFERENCE_HEADER
        template_row = 2 if sheet.max_row >= 2 else None
        for row in range(2, sheet.max_row + 1):
            for col in range(1, sheet.max_column + 1):
                sheet.cell(row, col).value = None
                sheet.row_dimensions[row].hidden = False
        ordered = sorted(records, key=lambda item: (item.project_code is None, item.project_code or 0, item.name.casefold(), item.major_version, item.document_date))
        highest = {item.normalized_name: max(other.major_version for other in records if other.normalized_name == item.normalized_name) for item in records}
        for row_number, record in enumerate(ordered, start=2):
            if template_row and row_number != template_row:
                _copy_row_style(sheet, template_row, row_number)
            values = (record.project_code, record.document_title, record.document_date, record.name, record.document_revision or None,
                      record.modified_by, record.major_version, record.smm_id or None, record.count or None, None, None)
            for col, value in enumerate(values, start=1):
                sheet.cell(row_number, col).value = value
            sheet.cell(row_number, 13).value = f"Duplicate with its version {highest[record.normalized_name]}" if "superseded_major_version" in record.reasons else None
            sheet.cell(row_number, OUTGOING_REFERENCE_COLUMN).value = outgoing_reference_without_revision(record.outgoing_reference, record.document_revision) or None
        for row in range(len(ordered) + 2, sheet.max_row + 1):
            for col in range(1, sheet.max_column + 1):
                sheet.cell(row, col).border = Border()
        sheet.auto_filter.ref = f"A1:P{max(1, len(ordered) + 1)}"
        sheet.auto_filter.filterColumn = []
        sheet.freeze_panes = "A2"
        for existing in workbook.worksheets:
            existing.sheet_view.tabSelected = False
        sheet.sheet_view.tabSelected = True
        workbook.active = workbook.index(sheet)
        workbook.save(temporary_path)
    finally:
        workbook.close()


def run_step1a(csv_path: Path, readiness_path: Path, workbook_path: Path, template_path: Path, backup_path: Path, start: date, end: date) -> dict[str, int]:
    if end < start:
        raise Step1aError("End date cannot be earlier than start date.")
    for label, path in (("CSV", csv_path), ("readiness workbook", readiness_path), ("ECM template", template_path)):
        if not path.is_file():
            raise FileNotFoundError(f"{label} was not found: {path}")
    workbook_path.parent.mkdir(parents=True, exist_ok=True)
    backup_path.parent.mkdir(parents=True, exist_ok=True)
    if workbook_path.resolve() == backup_path.resolve():
        raise Step1aError("Backup path must differ from the active ECM workbook.")
    if backup_path.exists():
        raise Step1aError(f"Run-specific backup already exists; refusing to overwrite it: {backup_path}")
    records = _read_records(csv_path, start, end)
    readiness = _read_readiness(readiness_path)
    apply_preliminary_rules(records, readiness)
    if workbook_path.exists():
        source = workbook_path
    else:
        source = template_path
    with tempfile.NamedTemporaryFile(prefix=f".{workbook_path.stem}-", suffix=workbook_path.suffix, dir=workbook_path.parent, delete=False) as temp:
        temporary_path = Path(temp.name)
    try:
        _build_updated_workbook(source, start, end, records, temporary_path)
        shutil.copy2(source, backup_path)
        os.replace(temporary_path, workbook_path)
    except Exception:
        temporary_path.unlink(missing_ok=True)
        raise
    result = {"rows_written": len(records), "pending_first_dates": sum(item.count == "" for item in records), "count_n": sum(item.count == "N" for item in records)}
    print(json.dumps({"event": "step1a_complete", **result}, sort_keys=True))
    print(f"Input CSV: {csv_path.resolve()}")
    print(f"Readiness workbook: {readiness_path.resolve()}")
    print(f"Updated ECM workbook: {workbook_path.resolve()}")
    print(f"Backup: {backup_path.resolve()}")
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description="Create the V3 ECM review worksheet from an exact-period CSV.")
    parser.add_argument("--start-date", required=True, type=date.fromisoformat)
    parser.add_argument("--end-date", required=True, type=date.fromisoformat)
    parser.add_argument("--csv", required=True, type=Path)
    parser.add_argument("--readiness-workbook", required=True, type=Path)
    parser.add_argument("--workbook", required=True, type=Path)
    parser.add_argument("--template", required=True, type=Path)
    parser.add_argument("--backup", required=True, type=Path)
    args = parser.parse_args()
    for label, path in (("CSV", args.csv), ("readiness workbook", args.readiness_workbook), ("ECM workbook", args.workbook), ("ECM template", args.template), ("backup", args.backup)):
        try:
            path.resolve().relative_to(CONTROL_PANEL_ROOT.resolve())
        except ValueError as error:
            raise Step1aError(f"{label} must remain under the V3 control-panel root: {path}") from error
    run_step1a(args.csv, args.readiness_workbook, args.workbook, args.template, args.backup, args.start_date, args.end_date)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (CsvValidationError, FileNotFoundError, OSError, Step1aError, ValueError) as error:
        print(f"Step 1a failed: {error}", file=sys.stderr)
        raise SystemExit(1)
