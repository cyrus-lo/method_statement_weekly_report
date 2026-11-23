"""Generate the independent V3 weekly interaction report.

The existing report implementation and SQL are imported read-only from the
legacy source tree. Credentials and all generated workbooks remain V3-owned.
"""

from __future__ import annotations

import argparse
import importlib.util
import os
import sys
import uuid
from contextlib import contextmanager
from datetime import date
from pathlib import Path
from types import ModuleType
from typing import Any, Iterator
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

CONTROL_PANEL_ROOT = Path(__file__).resolve().parents[1]
PROJECT_ROOT = CONTROL_PANEL_ROOT.parent
REFERENCE_REPORT_PATH = PROJECT_ROOT / "django" / "Method_Statement_Interaction_Weekly_Report.py"
DATABASE_ENVIRONMENT = {
    "PGHOST": "PGHOST",
    "PGDATABASE": "PGDATABASE",
    "PGUSER": "PGUSER",
    "PGPASSWORD": "PGPASSWORD",
}
OPTIONAL_DATABASE_ENVIRONMENT = ("PGPORT", "PGSSLMODE")
DEFAULT_REPORTING_TIMEZONE = "Asia/Hong_Kong"


class WeeklyReportError(ValueError):
    """Invalid V3 weekly-report configuration or output contract."""


def _v3_path(label: str, path: Path) -> Path:
    resolved = path.resolve()
    try:
        resolved.relative_to(CONTROL_PANEL_ROOT.resolve())
    except ValueError as error:
        raise WeeklyReportError(f"V3 {label} must remain under control_panel_v3: {resolved}") from error
    return resolved


def _load_reference_report() -> ModuleType:
    if not REFERENCE_REPORT_PATH.is_file():
        raise FileNotFoundError(f"Weekly report implementation was not found: {REFERENCE_REPORT_PATH}")
    spec = importlib.util.spec_from_file_location("v3_weekly_report_reference", REFERENCE_REPORT_PATH)
    if spec is None or spec.loader is None:
        raise WeeklyReportError("Could not load the weekly report implementation.")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    # The V3 builder supplies generic settings from django/.env explicitly;
    # never let the imported legacy module load dotenv a second time.
    module.load_dotenv = lambda *_args, **_kwargs: None
    return module


@contextmanager
def _v3_database_environment() -> Iterator[str]:
    required_values = {
        database_name: os.getenv(environment_name, "")
        for database_name, environment_name in DATABASE_ENVIRONMENT.items()
    }
    missing = [name for name, value in required_values.items() if not value.strip()]
    if missing:
        raise WeeklyReportError("Missing weekly database setting(s): " + ", ".join(missing))

    timezone_name = os.getenv("REPORTING_TIMEZONE", DEFAULT_REPORTING_TIMEZONE).strip()
    if not timezone_name:
        raise WeeklyReportError("REPORTING_TIMEZONE cannot be blank.")
    try:
        ZoneInfo(timezone_name)
    except (ValueError, ZoneInfoNotFoundError) as error:
        raise WeeklyReportError("REPORTING_TIMEZONE is not a recognized timezone.") from error

    mapped_values = dict(required_values)
    port_value = os.getenv("PGPORT", "5432").strip() or "5432"
    try:
        port = int(port_value)
    except ValueError as error:
        raise WeeklyReportError("PGPORT must be a whole number from 1 to 65535.") from error
    if not 1 <= port <= 65535:
        raise WeeklyReportError("PGPORT must be a whole number from 1 to 65535.")
    mapped_values["PGPORT"] = str(port)
    sslmode = os.getenv("PGSSLMODE", "").strip()
    if sslmode:
        mapped_values["PGSSLMODE"] = sslmode
    mapped_values["REPORTING_TIMEZONE"] = timezone_name

    managed_names = (*DATABASE_ENVIRONMENT, *OPTIONAL_DATABASE_ENVIRONMENT, "REPORTING_TIMEZONE")
    previous = {name: os.environ.get(name) for name in managed_names}
    try:
        for name in managed_names:
            os.environ.pop(name, None)
        os.environ.update(mapped_values)
        yield timezone_name
    finally:
        for name, value in previous.items():
            if value is None:
                os.environ.pop(name, None)
            else:
                os.environ[name] = value


def generate_weekly_report(
    start_date: date,
    end_date: date,
    template_path: Path,
    output_dir: Path,
    summary_sheet: str | None = None,
    *,
    reference_report: Any | None = None,
) -> Path:
    if end_date < start_date:
        raise WeeklyReportError("End date cannot be earlier than start date.")
    template_path = _v3_path("weekly-report template", template_path)
    output_dir = _v3_path("weekly-report output directory", output_dir)
    if not template_path.is_file():
        raise FileNotFoundError(f"V3 weekly-report template was not found: {template_path}")

    # The imported legacy script must not load a .env from its source tree.
    report = reference_report or _load_reference_report()
    for sql_path in (report.SESSIONS_SQL, report.SUBMISSION_COUNTS_SQL):
        if not sql_path.is_file():
            raise FileNotFoundError(f"Required weekly-report SQL file was not found: {sql_path}")

    output_name = (
        "Method_Statement_Interaction_Weekly_Report_"
        f"{start_date:%Y-%m-%d}_to_{end_date:%Y-%m-%d}.xlsx"
    )
    output_path = output_dir / output_name
    if output_path.exists():
        raise WeeklyReportError(f"V3 weekly-report output already exists; refusing to overwrite: {output_path}")

    stage_path = output_dir / f".{output_path.stem}.{uuid.uuid4().hex}.tmp.xlsx"
    try:
        with _v3_database_environment() as timezone_name:
            output_dir.mkdir(parents=True, exist_ok=True)
            # Also disable the reference module's legacy dotenv call when tests
            # inject an already-loaded module.
            report.load_dotenv = lambda *_args, **_kwargs: None
            parameters = (start_date, end_date)
            summary_title = report.build_summary_sheet_name(start_date, end_date)
            session_rows = report.sort_sessions(report.fetch_rows(report.SESSIONS_SQL, parameters))
            submission_rows = report.sort_submission_counts(
                report.fetch_rows(report.SUBMISSION_COUNTS_SQL, parameters)
            )
            summary_count = report.create_workbook(
                template_path,
                stage_path,
                summary_sheet,
                summary_title,
                session_rows,
                submission_rows,
            )
            if not stage_path.is_file():
                raise WeeklyReportError("Weekly report builder did not create the staged workbook.")
            os.replace(stage_path, output_path)
        print(
            "V3 weekly report created: "
            f"{output_path.resolve()} (period {start_date.isoformat()} to {end_date.isoformat()}, "
            f"timezone {timezone_name}, sessions {len(session_rows)}, "
            f"submission rows {len(submission_rows)}, summary projects {summary_count})"
        )
        return output_path
    finally:
        stage_path.unlink(missing_ok=True)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Generate a V3 weekly Method Statement interaction report.")
    parser.add_argument("--start-date", required=True, type=date.fromisoformat)
    parser.add_argument("--end-date", required=True, type=date.fromisoformat)
    parser.add_argument("--template", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--summary-sheet")
    args = parser.parse_args(argv)
    generate_weekly_report(
        args.start_date,
        args.end_date,
        args.template,
        args.output_dir,
        args.summary_sheet,
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"V3 weekly report failed: {type(error).__name__}: {error}", file=sys.stderr)
        raise SystemExit(1)
