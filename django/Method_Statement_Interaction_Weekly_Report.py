"""Create the Method Statement Interaction Weekly Report from PostgreSQL data."""

from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
from copy import copy
from datetime import date, datetime, time, timedelta, timezone, tzinfo
from pathlib import Path
from typing import Any, Iterable
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import psycopg
from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.styles import Color, Font, PatternFill
from openpyxl.utils import get_column_letter
from psycopg.rows import dict_row

ROOT = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = ROOT / "Method_Statement_Interaction_Weekly_Report_template.xlsx"
SESSIONS_SQL = ROOT / "sessions.sql"
SUBMISSION_COUNTS_SQL = ROOT / "submission_counts.sql"
REQUIRED_DATABASE_SETTINGS = ("PGHOST", "PGDATABASE", "PGUSER", "PGPASSWORD")
POSTGRES_GMT_OFFSET = re.compile(r"^GMT([+-])(\d{1,2})(?::?(\d{2}))?$")
DEFAULT_REPORTING_TIMEZONE = "Asia/Hong_Kong"

SUMMARY_MEASURES = {
    "Number of KMS User": "kms_user_count",
    "Number of Question Asked by KMS User to Retrieve Historical Method Statement Record": "count_retrieve_mode",
    "Number of PDF Viewed by KMS User": "count_pdf_viewed",
    "Number of Method Statement Content Selected by User to Further Edit in WORD Plugin": "count_content_copied",
    "Number of Question Asked by KMS User to Gather Work Trade Information": "count_ask_mode",
    "Number of Method Statement Drafted & Submitted to SMM": "count_submission_times",
    "Number of Reported Issue": "total_issue_count",
}
SESSIONS_HEADERS = (
    "project_code",
    "username",
    "created_at",
    "closed_at",
    "count_retrieve_mode",
    "count_pdf_viewed",
    "count_content_copied",
    "count_ask_mode",
    "count_submission_times",
    "total_issue_count",
    "pending_issue",
)
SUBMISSION_COUNTS_HEADERS = ("project_code", "count_submission_times")
SUMMARY_HEADER_ROW = 3
SUMMARY_FIRST_DATA_ROW = 4
# openpyxl writes Excel's internal width values, which are slightly larger than
# the width numbers shown in Excel's UI for this workbook/font combination.
SUMMARY_COLUMN_WIDTHS = {
    "A": 40.7265625,
    "B": 17.26953125,
    "C": 20.36328125,
    "D": 20.36328125,
    "E": 20.36328125,
    "F": 20.36328125,
    "G": 20.36328125,
    "H": 20.36328125,
    "I": 20.26953125,
}


def parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError as error:
        raise argparse.ArgumentTypeError("Dates must use YYYY-MM-DD format.") from error


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate the Method Statement Interaction Weekly Report."
    )
    parser.add_argument("--start-date", required=True, type=parse_date)
    parser.add_argument("--end-date", required=True, type=parse_date)
    parser.add_argument(
        "--template",
        type=Path,
        default=DEFAULT_TEMPLATE,
        help="Source workbook template (default: %(default)s).",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=ROOT / "output",
        help="Directory for generated workbooks (default: %(default)s).",
    )
    parser.add_argument(
        "--summary-sheet",
        help="Template worksheet to update. Defaults to the first worksheet.",
    )
    args = parser.parse_args()

    if args.end_date < args.start_date:
        parser.error("--end-date cannot be earlier than --start-date.")

    return args


def read_sql(path: Path) -> str:
    if not path.is_file():
        raise FileNotFoundError(f"Required SQL file was not found: {path}")
    return path.read_text(encoding="utf-8")


def build_summary_sheet_name(start_date: date, end_date: date) -> str:
    if start_date.month == end_date.month:
        return f"{start_date.day} - {end_date.day} {end_date.strftime('%b')} - Consolidated"
    return (
        f"{start_date.day} {start_date.strftime('%b')} - "
        f"{end_date.day} {end_date.strftime('%b')} - Consolidated"
    )


def get_database_settings() -> dict[str, str]:
    load_dotenv(ROOT / ".env")
    missing = [name for name in REQUIRED_DATABASE_SETTINGS if not os.getenv(name)]
    if missing:
        raise RuntimeError("Missing database settings in .env: " + ", ".join(missing))

    settings = {
        "host": os.environ["PGHOST"],
        "port": os.getenv("PGPORT", "5432"),
        "dbname": os.environ["PGDATABASE"],
        "user": os.environ["PGUSER"],
        "password": os.environ["PGPASSWORD"],
    }
    if sslmode := os.getenv("PGSSLMODE"):
        settings["sslmode"] = sslmode
    return settings


def get_reporting_timezone() -> str:
    timezone_name = os.getenv("REPORTING_TIMEZONE", DEFAULT_REPORTING_TIMEZONE).strip()
    if not timezone_name:
        raise RuntimeError("REPORTING_TIMEZONE cannot be blank.")
    try:
        ZoneInfo(timezone_name)
    except ZoneInfoNotFoundError as error:
        raise RuntimeError(f"Unknown REPORTING_TIMEZONE: {timezone_name}") from error
    return timezone_name


def set_connection_timezone(connection: psycopg.Connection[Any], timezone_name: str) -> None:
    """Set date boundaries and displayed timestamps to the reporting timezone."""
    with connection.cursor() as cursor:
        cursor.execute(
            "SELECT set_config('TimeZone', %s, false) AS timezone_name",
            (timezone_name,),
        )


def parse_postgres_timezone(value: str) -> tzinfo:
    if value.upper() == "UTC":
        return timezone.utc

    try:
        return ZoneInfo(value)
    except ZoneInfoNotFoundError:
        match = POSTGRES_GMT_OFFSET.match(value)
        if not match:
            return timezone.utc
        sign, hours_text, minutes_text = match.groups()
        offset = timedelta(hours=int(hours_text), minutes=int(minutes_text or 0))
        if sign == "+":
            offset = -offset
        return timezone(offset)


def get_connection_timezone(connection: psycopg.Connection[Any]) -> tzinfo:
    with connection.cursor() as cursor:
        cursor.execute("SELECT current_setting('TimeZone') AS timezone_name")
        return parse_postgres_timezone(str(cursor.fetchone()["timezone_name"]))


def normalise_row_values(row: dict[str, Any], target_timezone: tzinfo) -> dict[str, Any]:
    normalised = dict(row)
    for key, value in normalised.items():
        if isinstance(value, datetime) and value.tzinfo is not None:
            normalised[key] = value.astimezone(target_timezone).replace(tzinfo=None)
        elif isinstance(value, time) and value.tzinfo is not None:
            normalised[key] = value.replace(tzinfo=None)
    if "project_code" in normalised:
        normalised["project_code"] = normalise_project_code(normalised["project_code"])
    return normalised


def fetch_rows(sql_path: Path, parameters: tuple[date, date]) -> list[dict[str, Any]]:
    with psycopg.connect(**get_database_settings(), row_factory=dict_row) as connection:
        set_connection_timezone(connection, get_reporting_timezone())
        display_timezone = get_connection_timezone(connection)
        with connection.cursor() as cursor:
            cursor.execute(read_sql(sql_path), parameters)
            return [
                normalise_row_values(row, display_timezone)
                for row in cursor.fetchall()
            ]


def sort_sessions(rows: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    return sorted(
        rows,
        key=lambda row: (
            str(row.get("username") or "").casefold(),
            str(row.get("project_code") or "").casefold(),
        ),
    )


def sort_submission_counts(rows: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    return sorted(rows, key=lambda row: project_code_sort_key(row.get("project_code")))


def numeric_value(row: dict[str, Any], key: str) -> int:
    return int(row.get(key) or 0)


def project_code_sort_key(value: Any) -> tuple[int, Any]:
    project_code = normalise_project_code(value)
    if isinstance(project_code, int):
        return (0, project_code)
    if project_code is None:
        return (2, "")
    return (1, str(project_code).casefold())


def normalise_project_code(value: Any) -> Any:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    if text.isdigit():
        return int(text)
    return value


def canonical_project_code(value: Any) -> Any:
    project_code = normalise_project_code(value)
    if project_code in (1007, 1016):
        return 1016
    return project_code


def display_project_code(value: Any) -> Any:
    if canonical_project_code(value) == 1016:
        return "1016 (aka 1007)"
    return value


def include_project_code(value: Any) -> bool:
    project_code = normalise_project_code(value)
    return project_code is not None and project_code != 0


def excel_safe_value(value: Any) -> Any:
    if isinstance(value, str) and value.strip().isdigit():
        return int(value.strip())
    if isinstance(value, datetime) and value.tzinfo is not None:
        return value.replace(tzinfo=None)
    if isinstance(value, time) and value.tzinfo is not None:
        return value.replace(tzinfo=None)
    return value


def build_summary_rows(
    session_rows: Iterable[dict[str, Any]],
    submission_rows: Iterable[dict[str, Any]],
) -> list[dict[str, Any]]:
    summary: dict[Any, dict[str, Any]] = {}

    for row in session_rows:
        project_code = canonical_project_code(row.get("project_code"))
        if not include_project_code(project_code):
            continue
        project = summary.setdefault(
            project_code,
            {
                "project_code": project_code,
                "usernames": set(),
                "count_retrieve_mode": 0,
                "count_pdf_viewed": 0,
                "count_content_copied": 0,
                "count_ask_mode": 0,
                "count_submission_times": 0,
                "total_issue_count": 0,
            },
        )
        if username := row.get("username"):
            project["usernames"].add(username)
        for field in (
            "count_retrieve_mode",
            "count_pdf_viewed",
            "count_content_copied",
            "count_ask_mode",
            "total_issue_count",
        ):
            project[field] += numeric_value(row, field)

    for row in submission_rows:
        project_code = canonical_project_code(row.get("project_code"))
        if not include_project_code(project_code):
            continue
        project = summary.setdefault(
            project_code,
            {
                "project_code": project_code,
                "usernames": set(),
                "count_retrieve_mode": 0,
                "count_pdf_viewed": 0,
                "count_content_copied": 0,
                "count_ask_mode": 0,
                "count_submission_times": 0,
                "total_issue_count": 0,
            },
        )
        project["count_submission_times"] += numeric_value(row, "count_submission_times")

    result = []
    for project in summary.values():
        project["kms_user_count"] = len(project.pop("usernames"))
        result.append(project)
    return sorted(result, key=lambda row: project_code_sort_key(row["project_code"]))


def replace_sheet(workbook: Any, title: str) -> Any:
    if title in workbook.sheetnames:
        del workbook[title]
    return workbook.create_sheet(title)


def write_raw_sheet(
    workbook: Any,
    title: str,
    rows: list[dict[str, Any]],
    headers: tuple[str, ...],
) -> None:
    worksheet = replace_sheet(workbook, title)
    worksheet.append(headers)
    for cell in worksheet[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
    for row in rows:
        worksheet.append([excel_safe_value(row.get(header)) for header in headers])
    worksheet.auto_filter.ref = worksheet.dimensions
    worksheet.freeze_panes = "A2"
    set_column_widths(worksheet)


def normalise_heading(value: Any) -> str:
    return str(value).strip() if value is not None else ""


def find_summary_columns(worksheet: Any) -> dict[str, int]:
    columns = {
        normalise_heading(worksheet.cell(SUMMARY_HEADER_ROW, column).value): column
        for column in range(1, worksheet.max_column + 1)
    }
    missing = [heading for heading in SUMMARY_MEASURES if heading not in columns]
    if missing:
        raise ValueError("Summary heading(s) not found in row 3: " + "; ".join(missing))
    return {heading: columns[heading] for heading in SUMMARY_MEASURES}


def clear_summary_data(worksheet: Any) -> None:
    for row in worksheet.iter_rows(min_row=SUMMARY_FIRST_DATA_ROW, max_row=worksheet.max_row):
        for cell in row:
            cell.value = None


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


def force_black_summary_font(cell: Any) -> None:
    source_font = cell.font or Font()
    cell.font = Font(
        name=source_font.name,
        size=source_font.size,
        bold=source_font.bold,
        italic=source_font.italic,
        vertAlign=source_font.vertAlign,
        underline=source_font.underline,
        strike=source_font.strike,
        color="FF000000",
        charset=source_font.charset,
        outline=source_font.outline,
        shadow=source_font.shadow,
        condense=source_font.condense,
        extend=source_font.extend,
        family=source_font.family,
        scheme=source_font.scheme,
    )


def apply_summary_column_widths(worksheet: Any) -> None:
    for letter, width in SUMMARY_COLUMN_WIDTHS.items():
        worksheet.column_dimensions[letter].width = width


def write_summary_sheet(worksheet: Any, summary_rows: list[dict[str, Any]]) -> None:
    summary_columns = find_summary_columns(worksheet)
    style_row = SUMMARY_FIRST_DATA_ROW if worksheet.max_row >= SUMMARY_FIRST_DATA_ROW else None
    clear_summary_data(worksheet)

    for row_number, row in enumerate(summary_rows, start=SUMMARY_FIRST_DATA_ROW):
        if style_row is not None and row_number != style_row:
            copy_row_style(worksheet, style_row, row_number)
        project_cell = worksheet.cell(
            row_number,
            1,
            display_project_code(row["project_code"]),
        )
        force_black_summary_font(project_cell)
        for heading, key in SUMMARY_MEASURES.items():
            cell = worksheet.cell(row_number, summary_columns[heading], row[key])
            force_black_summary_font(cell)

    apply_summary_column_widths(worksheet)


def set_column_widths(worksheet: Any) -> None:
    for column in range(1, worksheet.max_column + 1):
        values = [worksheet.cell(row, column).value for row in range(1, worksheet.max_row + 1)]
        longest = max((len(str(value)) for value in values if value is not None), default=0)
        worksheet.column_dimensions[get_column_letter(column)].width = min(max(longest + 2, 12), 60)


def create_workbook(
    template_path: Path,
    output_path: Path,
    summary_sheet_name: str | None,
    summary_output_title: str,
    session_rows: list[dict[str, Any]],
    submission_rows: list[dict[str, Any]],
) -> int:
    if not template_path.is_file():
        raise FileNotFoundError(f"Workbook template was not found: {template_path}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_path, output_path)
    workbook = load_workbook(output_path)
    worksheet = workbook[summary_sheet_name] if summary_sheet_name else workbook.worksheets[0]
    worksheet.title = summary_output_title

    session_rows = [row for row in session_rows if include_project_code(row.get("project_code"))]
    submission_rows = [row for row in submission_rows if include_project_code(row.get("project_code"))]

    summary_rows = build_summary_rows(session_rows, submission_rows)
    write_summary_sheet(worksheet, summary_rows)
    write_raw_sheet(workbook, "Sessions", session_rows, SESSIONS_HEADERS)
    write_raw_sheet(
        workbook,
        "Submission Counts",
        submission_rows,
        SUBMISSION_COUNTS_HEADERS,
    )
    workbook.save(output_path)
    return len(summary_rows)


def main() -> int:
    args = parse_arguments()
    parameters = (args.start_date, args.end_date)
    summary_output_title = build_summary_sheet_name(args.start_date, args.end_date)
    session_rows = sort_sessions(fetch_rows(SESSIONS_SQL, parameters))
    submission_rows = sort_submission_counts(fetch_rows(SUBMISSION_COUNTS_SQL, parameters))
    output_name = (
        "Method_Statement_Interaction_Weekly_Report_"
        f"{args.start_date:%Y-%m-%d}_to_{args.end_date:%Y-%m-%d}.xlsx"
    )
    output_path = args.output_dir / output_name
    summary_count = create_workbook(
        args.template,
        output_path,
        args.summary_sheet,
        summary_output_title,
        session_rows,
        submission_rows,
    )

    print(f"Created: {output_path.resolve()}")
    print(f"Period: {args.start_date.isoformat()} to {args.end_date.isoformat()}")
    print(f"Reporting timezone: {get_reporting_timezone()}")
    print(f"Sessions: {len(session_rows)}")
    print(f"Submission Count rows: {len(submission_rows)}")
    print(f"Summary projects: {summary_count}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (FileNotFoundError, OSError, psycopg.Error, RuntimeError, ValueError) as error:
        print(f"Error: {error}", file=sys.stderr)
        raise SystemExit(1)
