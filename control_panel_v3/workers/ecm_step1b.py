"""Finalize Count (Y/N) on the V3 ECM workbook with staged, recoverable writes."""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from copy import copy
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.worksheet.filters import FilterColumn, Filters

from .ecm_step1a import OUTGOING_REFERENCE_HEADER, REVIEW_HEADERS, period_sheet_name

CONTROL_PANEL_ROOT = Path(__file__).resolve().parents[1]
FIRST_DATE_CUTOFF = date(2026, 4, 16)
COUNT_COLUMN = 9
SMM_ID_COLUMN = 8
DOCUMENT_REVISION_COLUMN = 5
REMARK_COLUMN = 13
OUTGOING_REFERENCE_COLUMN = 16
FIRST_VERSION_REMARK_PATTERN = re.compile(r"^1st version is submitted on (\d{1,2} [A-Za-z]{3} \d{4})$")
MISSING_SMM_ID_FILL = PatternFill("solid", fgColor="FFFFC000")


class Step1bError(ValueError):
    """Invalid Step 1b workbook data or V3 path contract."""


def validate_review_headers(worksheet: Any) -> None:
    headings = tuple(str(worksheet.cell(1, column).value or "").strip() for column in range(1, 12))
    if headings != REVIEW_HEADERS:
        raise Step1bError(f"Review worksheet headers A:K do not match the required ECM layout: {headings!r}.")
    if str(worksheet.cell(1, REMARK_COLUMN).value or "").strip() != "Remark":
        raise Step1bError("Review worksheet column M must have the header 'Remark'.")
    if str(worksheet.cell(1, OUTGOING_REFERENCE_COLUMN).value or "").strip() != OUTGOING_REFERENCE_HEADER:
        raise Step1bError(f"Review worksheet column P must have the header {OUTGOING_REFERENCE_HEADER!r}.")


def parse_manual_first_date(value: Any, row_number: int) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        text = value.strip()
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", text):
            try:
                return date.fromisoformat(text)
            except ValueError as error:
                raise Step1bError(f"Worksheet row {row_number}: M must contain a valid YYYY-MM-DD date.") from error
        saved = FIRST_VERSION_REMARK_PATTERN.fullmatch(text)
        if saved:
            try:
                return datetime.strptime(saved.group(1), "%d %b %Y").date()
            except ValueError as error:
                raise Step1bError(f"Worksheet row {row_number}: saved first-version date is invalid.") from error
        if text:
            raise Step1bError(f"Worksheet row {row_number}: M must be an Excel date, YYYY-MM-DD, or saved first-version remark; got {value!r}.")
    raise Step1bError(f"Worksheet row {row_number}: M requires a verified first document date for a pending Count row.")


def first_version_remark(first_date: date) -> str:
    return f"1st version is submitted on {first_date.day} {first_date:%b %Y}"


def populated_data_rows(worksheet: Any) -> list[int]:
    return [row for row in range(2, worksheet.max_row + 1) if any(worksheet.cell(row, column).value is not None for column in range(1, 9))]


def _prepare_finalized_workbook(source_path: Path, worksheet_name: str, staged_path: Path) -> tuple[int, int]:
    workbook = load_workbook(source_path)
    try:
        if worksheet_name not in workbook.sheetnames:
            raise Step1bError(f"Review worksheet {worksheet_name!r} was not found in {source_path.name}.")
        worksheet = workbook[worksheet_name]
        validate_review_headers(worksheet)
        data_rows = populated_data_rows(worksheet)
        decisions: list[tuple[int, str, date | None]] = []

        # Validate the entire worksheet before any durable workbook is touched.
        for row in data_rows:
            count_value = str(worksheet.cell(row, COUNT_COLUMN).value or "").strip().upper()
            if count_value not in {"", "Y", "N"}:
                raise Step1bError(f"Worksheet row {row}: Count (Y/N) must be blank, Y, or N; got {count_value!r}.")
            if count_value == "N":
                continue
            revision = str(worksheet.cell(row, DOCUMENT_REVISION_COLUMN).value or "").strip()
            if not revision:
                decisions.append((row, "Y", None))
                continue
            first_date = parse_manual_first_date(worksheet.cell(row, REMARK_COLUMN).value, row)
            final_count = "N" if first_date <= FIRST_DATE_CUTOFF else "Y"
            decisions.append((row, final_count, first_date))

        for row, final_count, first_date in decisions:
            worksheet.cell(row, COUNT_COLUMN).value = final_count
            if first_date is not None:
                worksheet.cell(row, REMARK_COLUMN).value = first_version_remark(first_date)
        for row in data_rows:
            if not str(worksheet.cell(row, COUNT_COLUMN).value or "").strip():
                worksheet.cell(row, COUNT_COLUMN).value = "Y"
            if not str(worksheet.cell(row, SMM_ID_COLUMN).value or "").strip():
                for column in range(1, worksheet.max_column + 1):
                    worksheet.cell(row, column).fill = copy(MISSING_SMM_ID_FILL)

        last_row = max(data_rows, default=1)
        worksheet.auto_filter.ref = f"A1:P{last_row}"
        worksheet.auto_filter.filterColumn = [FilterColumn(colId=COUNT_COLUMN - 1, filters=Filters(filter=["Y"]))]
        for sheet in workbook.worksheets:
            sheet.sheet_view.tabSelected = False
        worksheet.sheet_view.tabSelected = True
        workbook.active = workbook.index(worksheet)
        workbook.save(staged_path)
        return sum(count == "Y" for _, count, _ in decisions), sum(count == "N" for _, count, _ in decisions)
    finally:
        workbook.close()


def _powershell_filter_script(workbook_path: Path, worksheet_name: str) -> str:
    path = json.dumps(str(workbook_path.resolve()))
    sheet = json.dumps(worksheet_name)
    return f"""
$ErrorActionPreference = 'Stop'
$path = {path}
$sheetName = {sheet}
$excel = $null
$workbook = $null
$worksheet = $null
try {{
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {{ throw "Workbook not found: $path" }}
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $workbook = $excel.Workbooks.Open($path)
    $worksheet = $workbook.Worksheets.Item($sheetName)
    $worksheet.Activate()
    if ($worksheet.AutoFilterMode) {{ $worksheet.AutoFilterMode = $false }}
    $lastRow = $worksheet.Cells($worksheet.Rows.Count, 1).End(-4162).Row
    if ($lastRow -lt 1) {{ $lastRow = 1 }}
    $worksheet.Range("A1:P$lastRow").AutoFilter(9, "Y")
    $workbook.Save()
}} finally {{
    if ($null -ne $workbook) {{ try {{ $workbook.Close($false) }} catch {{}} }}
    if ($null -ne $worksheet) {{ try {{ [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) }} catch {{}} }}
    if ($null -ne $workbook) {{ try {{ [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }} catch {{}} }}
    if ($null -ne $excel) {{ try {{ $excel.Quit() }} catch {{}}; try {{ [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }} catch {{}} }}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}}
"""


def apply_count_filter_in_excel(workbook_path: Path, worksheet_name: str, *, timeout_seconds: int = 90) -> None:
    if sys.platform != "win32":
        raise RuntimeError("V3 Step 1b requires Windows Excel automation to apply the final Count filter.")
    lock_file = workbook_path.with_name(f"~${workbook_path.name}")
    if lock_file.exists():
        raise RuntimeError(f"Close the staged workbook before applying the Count filter: {workbook_path.name}")
    try:
        process = subprocess.Popen(
            ["powershell", "-NoProfile", "-NonInteractive", "-STA", "-ExecutionPolicy", "Bypass", "-Command", _powershell_filter_script(workbook_path, worksheet_name)],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding="utf-8", errors="replace",
        )
    except OSError as error:
        raise RuntimeError(f"Could not start Excel filter automation: {error}") from error
    try:
        stdout, stderr = process.communicate(timeout=timeout_seconds)
    except subprocess.TimeoutExpired as error:
        try:
            subprocess.run(["taskkill", "/PID", str(process.pid), "/T", "/F"], check=False, capture_output=True, text=True, timeout=10)
        except (OSError, subprocess.SubprocessError):
            process.kill()
        stdout, stderr = process.communicate()
        raise RuntimeError(f"Excel Count filter exceeded {timeout_seconds} seconds. {stderr or stdout}") from error
    if process.returncode != 0:
        detail = "\n".join(part.strip() for part in (stdout, stderr) if part and part.strip())
        raise RuntimeError(f"Excel could not apply the Count (Y/N) = Y filter. {detail or f'PowerShell exited with {process.returncode}.'}")


def _ensure_v3_path(label: str, path: Path) -> Path:
    resolved = path.resolve()
    try:
        resolved.relative_to(CONTROL_PANEL_ROOT.resolve())
    except ValueError as error:
        raise Step1bError(f"{label} must remain under the V3 control-panel root: {resolved}") from error
    return resolved


def _stage_copy(source: Path, target_directory: Path, suffix: str) -> Path:
    with tempfile.NamedTemporaryFile(prefix=".v3-step1b-", suffix=suffix, dir=target_directory, delete=False) as stream:
        return Path(stream.name)


def run_step1b(
    workbook_path: Path,
    worksheet_name: str,
    backup_path: Path,
    post_backup_path: Path,
    *,
    excel_filter: Callable[[Path, str], None] = apply_count_filter_in_excel,
    excel_filter_timeout_seconds: int = 90,
) -> tuple[int, int]:
    workbook_path = _ensure_v3_path("ECM workbook", workbook_path)
    backup_path = _ensure_v3_path("Pre-Step-1b backup", backup_path)
    post_backup_path = _ensure_v3_path("Post-Step-1b snapshot", post_backup_path)
    if not workbook_path.is_file():
        raise FileNotFoundError(f"V3 ECM review workbook was not found: {workbook_path}")
    if len({workbook_path, backup_path, post_backup_path}) != 3:
        raise Step1bError("Workbook, pre-backup, and post-backup paths must all be different.")
    if backup_path.exists() or post_backup_path.exists():
        raise Step1bError("Run-specific Step 1b backup or post-snapshot already exists; refusing to overwrite evidence.")
    if workbook_path.with_name(f"~${workbook_path.name}").exists():
        raise Step1bError(f"Close the V3 ECM workbook before running Step 1b: {workbook_path.name}")
    backup_path.parent.mkdir(parents=True, exist_ok=True)
    post_backup_path.parent.mkdir(parents=True, exist_ok=True)
    staged_path = _stage_copy(workbook_path, workbook_path.parent, ".xlsx")
    post_staged_path: Path | None = None
    try:
        counts = _prepare_finalized_workbook(workbook_path, worksheet_name, staged_path)
        # Keep the prior workbook as the recovery point before running Excel on the staged copy.
        shutil.copy2(workbook_path, backup_path)
        excel_filter_applied = False
        if excel_filter is apply_count_filter_in_excel:
            excel_filter(staged_path, worksheet_name, timeout_seconds=excel_filter_timeout_seconds)
            excel_filter_applied = True
        else:
            excel_filter(staged_path, worksheet_name)
        if not excel_filter_applied:
            filtered_path = _stage_copy(staged_path, staged_path.parent, staged_path.suffix)
            try:
                workbook = load_workbook(staged_path)
                try:
                    sheet = workbook[worksheet_name]
                    sheet.auto_filter.ref = f"A1:P{max(1, sheet.max_row)}"
                    sheet.auto_filter.filterColumn = [FilterColumn(colId=COUNT_COLUMN - 1, filters=Filters(filter=["Y"]))]
                    workbook.save(filtered_path)
                finally:
                    workbook.close()
                os.replace(filtered_path, staged_path)
            finally:
                filtered_path.unlink(missing_ok=True)
        post_staged_path = _stage_copy(staged_path, post_backup_path.parent, post_backup_path.suffix)
        shutil.copy2(staged_path, post_staged_path)
        post_committed = False
        try:
            os.replace(post_staged_path, post_backup_path)
            post_staged_path = None
            post_committed = True
            os.replace(staged_path, workbook_path)
        except Exception:
            # If committing the post snapshot fails after the active workbook changed, restore the original atomically.
            if post_committed and post_backup_path.exists():
                post_backup_path.unlink()
            rollback_path = _stage_copy(backup_path, workbook_path.parent, workbook_path.suffix)
            try:
                shutil.copy2(backup_path, rollback_path)
                os.replace(rollback_path, workbook_path)
            finally:
                rollback_path.unlink(missing_ok=True)
            raise
        print(json.dumps({"event": "step1b_complete", "count_y": counts[0], "count_n": counts[1], "worksheet": worksheet_name}, sort_keys=True))
        print(f"Updated ECM workbook: {workbook_path}")
        print(f"Pre-Step-1b backup: {backup_path}")
        print(f"Post-Step-1b snapshot: {post_backup_path}")
        return counts
    finally:
        staged_path.unlink(missing_ok=True)
        if post_staged_path is not None:
            post_staged_path.unlink(missing_ok=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="Finalize V3 ECM review decisions after first-date review.")
    parser.add_argument("--start-date", required=True, type=date.fromisoformat)
    parser.add_argument("--end-date", required=True, type=date.fromisoformat)
    parser.add_argument("--workbook", required=True, type=Path)
    parser.add_argument("--worksheet")
    parser.add_argument("--backup", required=True, type=Path)
    parser.add_argument("--post-backup", required=True, type=Path)
    parser.add_argument("--excel-filter-timeout-seconds", type=int, default=90)
    args = parser.parse_args()
    if args.end_date < args.start_date:
        parser.error("--end-date cannot be earlier than --start-date")
    for label, path in (("ECM workbook", args.workbook), ("pre-backup", args.backup), ("post-backup", args.post_backup)):
        _ensure_v3_path(label, path)
    run_step1b(args.workbook, args.worksheet or period_sheet_name(args.start_date, args.end_date), args.backup, args.post_backup, excel_filter_timeout_seconds=args.excel_filter_timeout_seconds)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, RuntimeError, Step1bError, ValueError) as error:
        print(f"V3 Step 1b failed: {error}", file=sys.stderr)
        raise SystemExit(1)
