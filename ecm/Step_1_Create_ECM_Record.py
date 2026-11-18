"""Update the ECM review workbook from an ECM CSV export and readiness data.

The persistent ECM_Record.xlsx workbook is updated in place. Before each update,
the existing workbook is copied to a dated backup. The actual ECM first-document-
date search is intentionally supplied through a local CSV until its ECM search
procedure is specified.
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import shutil
import subprocess
import sys
from copy import copy
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.worksheet.filters import FilterColumn, Filters

ROOT = Path(__file__).resolve().parent
CSV_DIRECTORY = ROOT / "csv_from_ecm"
DEFAULT_ECM_TEMPLATE = ROOT / "ECM_Record_template.xlsx"
DEFAULT_SMM_WORKBOOK = ROOT / "SMM_KMS_PM_Master_Document_template.xlsx"
DEFAULT_OUTPUT_DIRECTORY = ROOT / "output"
DEFAULT_ECM_RECORD = DEFAULT_OUTPUT_DIRECTORY / "ECM_Record.xlsx"
ECM_SHEET_NAME = "ECM_1-12 Jul (WIP)"
ECM_FALLBACK_SHEET_NAME = "ECM_1-12 Jul (OK)"
ECM_TEMPLATE_SHEET_NAMES = ("ECM_template", ECM_SHEET_NAME, ECM_FALLBACK_SHEET_NAME)
READINESS_SHEET_NAME = "Project Realiness Chris-In Use"
FIRST_DATE_CUTOFF = date(2026, 4, 16)
EXCEL_FILTER_TIMEOUT_SECONDS = 90
MISSING_SMM_ID_FILL = PatternFill("solid", fgColor="FFFFC000")

CSV_HEADERS = (
    "Document title",
    "Document date",
    "Name",
    "Document revision",
    "Modified By",
    "Major Version",
    "SMM ID",
    "Modified On",
    "Outgoing ref. no.",
    "Path",
    "Link",
)
REVIEW_HEADERS = (
    "Project Code",
    "Document title",
    "Document date",
    "Name",
    "Document revision",
    "Modified By",
    "Major Version",
    "SMM ID",
    "Count (Y/N)",
    "Use SMM (Y/N)",
    "Use KMS (Y/N)",
)
PROJECT_CODE_PATTERN = re.compile(r"^[^/]*/\s*(\d+)\s*/")
IDENTIFIER_TOKEN_PATTERN = re.compile(r"[A-Za-z0-9]+")
FILE_EXTENSION_PATTERN = re.compile(r"\.(?:pdf|docx?|xlsx?|xlsm)$", re.IGNORECASE)
PROJECT_FROM_READINESS_PATTERN = re.compile(r"^\s*(\d+)\b")


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
    path: str
    link: str
    project_code: int | None = None
    count: str = "N"
    reasons: list[str] = field(default_factory=list)

    @property
    def normalized_name(self) -> str:
        return self.name.strip().casefold()


class InputValidationError(ValueError):
    """Raised when a supplied CSV or workbook does not meet the required contract."""


def parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError as error:
        raise argparse.ArgumentTypeError("Dates must use YYYY-MM-DD format.") from error


def parse_ecm_datetime(value: str, field_name: str, source_row: int) -> datetime:
    value = value.strip()
    for date_format in ("%m/%d/%Y, %I:%M %p", "%m/%d/%Y %I:%M %p"):
        try:
            return datetime.strptime(value, date_format)
        except ValueError:
            continue
    raise InputValidationError(
        f"CSV row {source_row}: {field_name!r} must use M/D/YYYY, H:MM AM/PM; got {value!r}."
    )


def parse_optional_ecm_datetime(value: str, field_name: str, source_row: int) -> datetime | None:
    return None if not value.strip() else parse_ecm_datetime(value, field_name, source_row)


def normalise_text(value: Any) -> str:
    return str(value or "").strip()


def normalise_project_code(value: Any) -> int | None:
    if value is None:
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float) and value.is_integer():
        return int(value)
    match = PROJECT_FROM_READINESS_PATTERN.match(normalise_text(value))
    return int(match.group(1)) if match else None


def project_code_from_delimited_value(value: str) -> int | None:
    match = PROJECT_CODE_PATTERN.match(value.strip())
    return int(match.group(1)) if match else None


def comparable_identifier_parts(value: str) -> tuple[str, ...]:
    value_without_extension = FILE_EXTENSION_PATTERN.sub("", value.strip())
    return tuple(
        part.casefold()
        for part in IDENTIFIER_TOKEN_PATTERN.findall(value_without_extension)
    )


def project_code_from_matching_identifiers(name: str, outgoing_reference: str) -> int | None:
    name_parts = comparable_identifier_parts(name)
    reference_parts = comparable_identifier_parts(outgoing_reference)
    if (
        len(name_parts) >= 2
        and name_parts == reference_parts
        and name_parts[1].isdigit()
    ):
        return int(name_parts[1])
    return None


def project_code_from_fields(
    name: str,
    outgoing_reference: str,
    source_row: int,
) -> int | None:
    project_code = project_code_from_delimited_value(name)
    if project_code is not None:
        return project_code

    project_code = project_code_from_delimited_value(outgoing_reference)
    if project_code is not None:
        print(
            f"WARNING: CSV row {source_row}: project code {project_code} "
            "was recovered from Outgoing ref. no. because Name did not match.",
            file=sys.stderr,
        )
        return project_code

    project_code = project_code_from_matching_identifiers(name, outgoing_reference)
    if project_code is not None:
        print(
            f"WARNING: CSV row {source_row}: project code {project_code} "
            "was inferred by matching Name and Outgoing ref. no. after "
            "ignoring punctuation and file extensions. Please verify.",
            file=sys.stderr,
        )
    return project_code


def canonical_project_code(project_code: int | None) -> int | None:
    return 1016 if project_code == 1007 else project_code


def find_newest_csv(csv_directory: Path) -> Path:
    candidates = [
        path
        for path in csv_directory.glob("properties*.csv")
        if path.is_file() and path.name.casefold() != "properties*.csv"
    ]
    if not candidates:
        raise FileNotFoundError(f"No properties*.csv input was found in {csv_directory}")
    return max(candidates, key=lambda path: path.stat().st_mtime)


def read_ecm_records(csv_path: Path, start_date: date, end_date: date) -> list[EcmRecord]:
    if not csv_path.is_file():
        raise FileNotFoundError(f"ECM CSV was not found: {csv_path}")

    with csv_path.open("r", encoding="utf-8-sig", newline="") as source:
        reader = csv.DictReader(source)
        if reader.fieldnames is None:
            raise InputValidationError("The ECM CSV does not contain a header row.")
        headers = [normalise_text(header) for header in reader.fieldnames]
        missing_headers = [header for header in CSV_HEADERS if header not in headers]
        if missing_headers:
            raise InputValidationError("ECM CSV is missing header(s): " + ", ".join(missing_headers))

        records: list[EcmRecord] = []
        for source_row, raw_row in enumerate(reader, start=2):
            row = {normalise_text(key): normalise_text(value) for key, value in raw_row.items()}
            document_date = parse_ecm_datetime(row["Document date"], "Document date", source_row)
            if not start_date <= document_date.date() <= end_date:
                continue
            try:
                major_version = int(row["Major Version"])
            except ValueError as error:
                raise InputValidationError(
                    f"CSV row {source_row}: Major Version must be an integer; got {row['Major Version']!r}."
                ) from error
            records.append(
                EcmRecord(
                    source_row=source_row,
                    document_title=row["Document title"],
                    document_date=document_date,
                    name=row["Name"],
                    document_revision=row["Document revision"],
                    modified_by=row["Modified By"],
                    major_version=major_version,
                    smm_id=row["SMM ID"],
                    modified_on=parse_optional_ecm_datetime(row["Modified On"], "Modified On", source_row),
                    outgoing_reference=row["Outgoing ref. no."],
                    path=row["Path"],
                    link=row["Link"],
                    project_code=project_code_from_fields(
                        row["Name"],
                        row["Outgoing ref. no."],
                        source_row,
                    ),
                )
            )
    return records


def read_readiness(smm_workbook_path: Path) -> dict[int, bool]:
    if not smm_workbook_path.is_file():
        raise FileNotFoundError(f"SMM workbook was not found: {smm_workbook_path}")

    workbook = load_workbook(smm_workbook_path, read_only=True, data_only=True)
    try:
        if READINESS_SHEET_NAME not in workbook.sheetnames:
            raise InputValidationError(
                f"Readiness sheet {READINESS_SHEET_NAME!r} was not found in {smm_workbook_path.name}."
            )
        worksheet = workbook[READINESS_SHEET_NAME]
        readiness: dict[int, bool] = {}
        for row_number in range(2, worksheet.max_row + 1):
            project_code = normalise_project_code(worksheet.cell(row_number, 3).value)
            if project_code is None:
                continue
            readiness_text = normalise_text(worksheet.cell(row_number, 5).value).casefold()
            if project_code in readiness:
                raise InputValidationError(
                    f"Duplicate project code {project_code} in readiness sheet row {row_number}."
                )
            readiness[project_code] = readiness_text not in {"no", "n"}
        if not readiness:
            raise InputValidationError("No project codes were found in column C of the readiness sheet.")
        return readiness
    finally:
        workbook.close()


def read_first_dates(first_date_path: Path | None) -> dict[str, date]:
    """Read a temporary first-date export until the ECM search procedure is implemented.

    The file must contain `Name` and `First Document Date` headers. A later ECM
    lookup implementation can replace this function without changing the review
    workbook rules.
    """
    if first_date_path is None:
        raise InputValidationError(
            "Final Step 1 processing requires --first-date-file until the ECM search procedure is implemented. "
            "The file must contain Name and First Document Date columns."
        )
    if not first_date_path.is_file():
        raise FileNotFoundError(f"First-date file was not found: {first_date_path}")

    with first_date_path.open("r", encoding="utf-8-sig", newline="") as source:
        reader = csv.DictReader(source)
        if reader.fieldnames is None:
            raise InputValidationError("The first-date file does not contain a header row.")
        headers = {normalise_text(header) for header in reader.fieldnames}
        required_headers = {"Name", "First Document Date"}
        if not required_headers.issubset(headers):
            raise InputValidationError("First-date file must contain Name and First Document Date headers.")

        first_dates: dict[str, date] = {}
        for row_number, raw_row in enumerate(reader, start=2):
            row = {normalise_text(key): normalise_text(value) for key, value in raw_row.items()}
            normalized_name = row["Name"].casefold()
            if not normalized_name:
                raise InputValidationError(f"First-date file row {row_number}: Name cannot be blank.")
            try:
                first_date = date.fromisoformat(row["First Document Date"])
            except ValueError as error:
                raise InputValidationError(
                    f"First-date file row {row_number}: First Document Date must use YYYY-MM-DD."
                ) from error
            existing = first_dates.get(normalized_name)
            if existing is not None and existing != first_date:
                raise InputValidationError(
                    f"First-date file has conflicting first dates for {row['Name']!r}."
                )
            first_dates[normalized_name] = first_date
        return first_dates


def apply_counting_rules(records: list[EcmRecord], readiness: dict[int, bool], first_dates: dict[str, date]) -> None:
    highest_versions: dict[str, int] = {}
    for record in records:
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
        if len(candidates) > 1:
            canonical = max(
                candidates,
                key=lambda record: (record.modified_on or datetime.min, -record.source_row),
            )
            for record in candidates:
                if record is not canonical:
                    record.reasons.append("duplicate_max_major_version")

    for record in records:
        first_date = first_dates.get(record.normalized_name)
        if first_date is None:
            record.reasons.append("first_document_not_found")
        elif first_date < FIRST_DATE_CUTOFF:
            record.reasons.append("first_document_before_cutoff")
        record.count = "Y" if not record.reasons else "N"


def period_sheet_name(start_date: date, end_date: date) -> str:
    if start_date.month == end_date.month:
        return f"ECM_{start_date.day}-{end_date.day} {end_date:%b} (WIP)"
    return f"ECM_{start_date.day} {start_date:%b}-{end_date.day} {end_date:%b} (WIP)"


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


def validate_review_headers(worksheet: Any) -> None:
    headings = tuple(normalise_text(worksheet.cell(1, column).value) for column in range(1, 12))
    if headings != REVIEW_HEADERS:
        raise InputValidationError(
            f"Review worksheet headers A:K do not match the required ECM layout: {headings!r}."
        )


def clear_review_rows(worksheet: Any) -> None:
    for row_number in range(2, worksheet.max_row + 1):
        for column in range(1, worksheet.max_column + 1):
            worksheet.cell(row_number, column).value = None


def get_template_review_worksheet(workbook: Any, workbook_name: str) -> Any:
    """Return the preferred fixed review template worksheet."""
    for worksheet_name in ECM_TEMPLATE_SHEET_NAMES:
        if worksheet_name in workbook.sheetnames:
            return workbook[worksheet_name]
    expected_names = ", ".join(repr(name) for name in ECM_TEMPLATE_SHEET_NAMES)
    raise InputValidationError(
        f"Review worksheet must be one of {expected_names} in {workbook_name}."
    )


def write_review_workbook(
    workbook_path: Path,
    start_date: date,
    end_date: date,
    records: list[EcmRecord],
) -> None:
    if not workbook_path.is_file():
        raise FileNotFoundError(f"ECM review workbook was not found: {workbook_path}")

    workbook = load_workbook(workbook_path)
    try:
        source_worksheet = get_template_review_worksheet(workbook, workbook_path.name)
        validate_review_headers(source_worksheet)
        target_sheet_name = period_sheet_name(start_date, end_date)
        if target_sheet_name in workbook.sheetnames:
            del workbook[target_sheet_name]
        worksheet = workbook.copy_worksheet(source_worksheet)
        worksheet.title = target_sheet_name
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
        for row_number, record in enumerate(ordered_records, start=2):
            if style_row is not None and row_number != style_row:
                copy_row_style(worksheet, style_row, row_number)
            worksheet.row_dimensions[row_number].hidden = False
            values = (
                record.project_code,
                record.document_title,
                record.document_date,
                record.name,
                record.document_revision or None,
                record.modified_by,
                record.major_version,
                record.smm_id or None,
                record.count,
                None,
                None,
            )
            for column, value in enumerate(values, start=1):
                worksheet.cell(row_number, column).value = value

        for row_number, record in enumerate(ordered_records, start=2):
            if not record.smm_id:
                for column in range(1, worksheet.max_column + 1):
                    worksheet.cell(row_number, column).fill = copy(MISSING_SMM_ID_FILL)

        for row_number in range(1, worksheet.max_row + 1):
            worksheet.row_dimensions[row_number].hidden = False
        worksheet.auto_filter.ref = f"A1:O{max(1, len(ordered_records) + 1)}"
        worksheet.auto_filter.filterColumn = [
            FilterColumn(colId=8, filters=Filters(filter=["Y"]))
        ]
        worksheet.freeze_panes = "A2"
        for sheet in workbook.worksheets:
            sheet.sheet_view.tabSelected = False
        worksheet.sheet_view.tabSelected = True
        workbook.active = workbook.index(worksheet)
        workbook.save(workbook_path)
    finally:
        workbook.close()


def _count_filter_powershell_script(output_path: Path, worksheet_name: str) -> str:
    """Build the one-session PowerShell script used to apply the Count filter."""
    script = f"""
$ErrorActionPreference = 'Stop'
$path = {json.dumps(str(output_path.resolve()))}
$sheetName = {json.dumps(worksheet_name)}
$excel = $null
$workbook = $null
$worksheet = $null
try {{
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {{
        throw "Workbook was not found: $path"
    }}
    $excel = New-Object -ComObject Excel.Application
    if ($null -eq $excel) {{
        throw "Excel.Application could not be created."
    }}
    Write-Output "Excel session: started"
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $workbook = $excel.Workbooks.Open($path)
    if ($null -eq $workbook) {{
        throw "Excel returned no workbook when opening: $path"
    }}
    Write-Output "Excel session: workbook opened: $path"
    $availableSheets = @()
    for ($attempt = 1; $attempt -le 10 -and $null -eq $worksheet; $attempt++) {{
        try {{
            $worksheets = $workbook.Worksheets
            if ($null -ne $worksheets) {{
                $worksheet = $worksheets.Item($sheetName)
                if ($attempt -eq 1) {{
                    $availableSheets = @(
                        1..$worksheets.Count |
                        ForEach-Object {{ $worksheets.Item($_).Name }}
                    )
                }}
            }}
        }} catch {{
            if ($attempt -eq 10) {{ throw }}
        }}
        if ($null -eq $worksheet) {{ Start-Sleep -Milliseconds 500 }}
    }}
    if ($null -eq $worksheet) {{
        $sheetList = if ($availableSheets.Count -gt 0) {{ $availableSheets -join ', ' }} else {{ '<unavailable>' }}
        throw "Worksheet was not available after waiting: $sheetName. Available sheets: $sheetList"
    }}
    $worksheet.Activate()
    if ($worksheet.AutoFilterMode) {{ $worksheet.AutoFilterMode = $false }}
    $lastRow = $worksheet.Cells($worksheet.Rows.Count, 1).End(-4162).Row
    $worksheet.Range("A1:O$lastRow").AutoFilter(9, "Y")
    $workbook.Save()
    Write-Output "Excel session: Count (Y/N) = Y filter applied and workbook saved"
}} finally {{
    if ($null -ne $workbook) {{
        try {{
            $workbook.Close($false)
            Write-Output "Excel session: workbook closed"
        }} catch {{
            Write-Warning "Excel session: workbook close failed: $($_.Exception.Message)"
        }}
    }}
    if ($null -ne $worksheet) {{
        try {{ [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) }} catch {{}}
    }}
    if ($null -ne $workbook) {{
        try {{ [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }} catch {{}}
    }}
    if ($null -ne $excel) {{
        try {{
            $excel.Quit()
            Write-Output "Excel session: application quit"
        }} catch {{
            Write-Warning "Excel session: application quit failed: $($_.Exception.Message)"
        }}
        try {{ [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }} catch {{}}
    }}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}}
"""
    return script


def apply_count_filter_in_excel(output_path: Path, worksheet_name: str) -> None:
    """Apply Count=Y through one hidden Excel session, then close that session."""
    if sys.platform != "win32":
        return
    if not output_path.is_file():
        raise RuntimeError(f"Excel filter workbook was not found: {output_path}")
    lock_file = output_path.with_name(f"~${output_path.name}")
    if lock_file.exists():
        raise RuntimeError(f"Close the workbook in Excel before applying the filter: {output_path.name}")

    script = _count_filter_powershell_script(output_path, worksheet_name)
    try:
        process = subprocess.Popen(
            [
                "powershell",
                "-NoProfile",
                "-NonInteractive",
                "-STA",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                script,
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
    except FileNotFoundError as error:
        details = str(error)
        raise RuntimeError(
            "Excel could not apply the Count (Y/N) = Y filter. "
            f"Workbook: {output_path.resolve()}; worksheet: {worksheet_name!r}. "
            "Close the generated workbook and ensure desktop Excel is available. "
            f"Details: {details}"
        ) from error

    try:
        stdout, stderr = process.communicate(timeout=EXCEL_FILTER_TIMEOUT_SECONDS)
    except subprocess.TimeoutExpired as error:
        try:
            subprocess.run(
                ["taskkill", "/PID", str(process.pid), "/T", "/F"],
                check=False,
                capture_output=True,
                text=True,
                timeout=10,
            )
        except (FileNotFoundError, subprocess.SubprocessError):
            try:
                process.kill()
            except OSError:
                pass
        try:
            stdout, stderr = process.communicate(timeout=10)
        except subprocess.TimeoutExpired:
            process.kill()
            stdout, stderr = process.communicate()
        details = f"Excel session exceeded the {EXCEL_FILTER_TIMEOUT_SECONDS}-second timeout: {error}"
        raise RuntimeError(
            "Excel could not apply the Count (Y/N) = Y filter. "
            f"Workbook: {output_path.resolve()}; worksheet: {worksheet_name!r}. "
            "Close the generated workbook and ensure desktop Excel is available. "
            f"Details: {details}"
        ) from error

    if process.returncode != 0:
        details = "\n".join(
            part.strip()
            for part in (stdout or "", stderr or "")
            if part and part.strip()
        ) or f"PowerShell exited with code {process.returncode}."
        error = subprocess.CalledProcessError(process.returncode, process.args, output=stdout, stderr=stderr)
        raise RuntimeError(
            "Excel could not apply the Count (Y/N) = Y filter. "
            f"Workbook: {output_path.resolve()}; worksheet: {worksheet_name!r}. "
            "Close the generated workbook and ensure desktop Excel is available. "
            f"Details: {details}"
        ) from error


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Step 1: update ECM_Record.xlsx and create a dated backup.")
    parser.add_argument("--start-date", required=True, type=parse_date)
    parser.add_argument("--end-date", required=True, type=parse_date)
    parser.add_argument("--csv", type=Path, help="ECM properties CSV. Defaults to the newest properties*.csv.")
    parser.add_argument("--smm-workbook", type=Path, default=DEFAULT_SMM_WORKBOOK)
    parser.add_argument(
        "--ecm-record",
        type=Path,
        default=DEFAULT_ECM_RECORD,
        help="Persistent ECM workbook to update in place.",
    )
    parser.add_argument(
        "--ecm-template",
        type=Path,
        default=DEFAULT_ECM_TEMPLATE,
        help="Template used to initialize --ecm-record on the first run.",
    )
    parser.add_argument(
        "--first-date-file",
        type=Path,
        help="Temporary CSV with Name and First Document Date columns, pending ECM search integration.",
    )
    parser.add_argument("--output", type=Path, help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.end_date < args.start_date:
        parser.error("--end-date cannot be earlier than --start-date.")
    return args


def main() -> int:
    args = parse_arguments()
    csv_path = args.csv or find_newest_csv(CSV_DIRECTORY)
    ecm_record_path = args.ecm_record
    backup_path = args.output or (
        DEFAULT_OUTPUT_DIRECTORY / f"ECM_Record_{args.start_date:%Y-%m-%d}_to_{args.end_date:%Y-%m-%d}.xlsx"
    )
    ecm_record_path.parent.mkdir(parents=True, exist_ok=True)
    if not ecm_record_path.is_file():
        if not args.ecm_template.is_file():
            raise FileNotFoundError(f"ECM review template was not found: {args.ecm_template}")
        shutil.copy2(args.ecm_template, ecm_record_path)
    if backup_path.resolve() == ecm_record_path.resolve():
        raise InputValidationError("The ECM backup path must differ from the ECM_Record.xlsx path.")
    shutil.copy2(ecm_record_path, backup_path)
    records = read_ecm_records(csv_path, args.start_date, args.end_date)
    readiness = read_readiness(args.smm_workbook)
    first_dates = read_first_dates(args.first_date_file)
    apply_counting_rules(records, readiness, first_dates)
    write_review_workbook(ecm_record_path, args.start_date, args.end_date, records)
    apply_count_filter_in_excel(ecm_record_path, period_sheet_name(args.start_date, args.end_date))

    print(f"Input CSV: {csv_path.resolve()}")
    print(f"Readiness workbook: {args.smm_workbook.resolve()}")
    print(f"Updated ECM record: {ecm_record_path.resolve()}")
    print(f"Backup: {backup_path.resolve()}")
    print(f"Rows written: {len(records)}")
    print(f"Count Y: {sum(record.count == 'Y' for record in records)}")
    print(f"Count N: {sum(record.count == 'N' for record in records)}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (FileNotFoundError, InputValidationError, OSError, RuntimeError, ValueError) as error:
        print(f"Error: {error}", file=sys.stderr)
        raise SystemExit(1)
