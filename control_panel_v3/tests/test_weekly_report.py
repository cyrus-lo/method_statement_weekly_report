from __future__ import annotations

import hashlib
import os
import shutil
from datetime import date
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from control_panel_v3.services.config import CONTROL_PANEL_ROOT, load_settings
from control_panel_v3.services.workflow_builders import PreflightError, prepare_workflow
from control_panel_v3.services.workflows import REPORT_WORKFLOW_IDS, ecm_workflow_gate_error
from control_panel_v3.workers.weekly_report import WeeklyReportError, generate_weekly_report


REPORTING_ENV = {
    "CONTROL_PANEL_V3_WEEKLY_PGHOST": "v3-db.example.test",
    "CONTROL_PANEL_V3_WEEKLY_PGDATABASE": "weekly_test",
    "CONTROL_PANEL_V3_WEEKLY_PGUSER": "read_only_test",
    "CONTROL_PANEL_V3_WEEKLY_PGPASSWORD": "test-password-not-a-real-secret",
    "CONTROL_PANEL_V3_WEEKLY_PGPORT": "5433",
    "CONTROL_PANEL_V3_WEEKLY_PGSSLMODE": "require",
    "CONTROL_PANEL_V3_WEEKLY_REPORTING_TIMEZONE": "Asia/Hong_Kong",
}


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _make_template(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Consolidated"
    sheet["A1"] = "Template title"
    sheet["A2"] = "Template subtitle"
    sheet["A3"] = "Project"
    for column, heading in enumerate((
        "Number of KMS User",
        "Number of Question Asked by KMS User to Retrieve Historical Method Statement Record",
        "Number of PDF Viewed by KMS User",
        "Number of Method Statement Content Selected by User to Further Edit in WORD Plugin",
        "Number of Question Asked by KMS User to Gather Work Trade Information",
        "Number of Method Statement Drafted & Submitted to SMM",
        "Number of Reported Issue",
    ), start=2):
        sheet.cell(3, column).value = heading
    for column in range(1, 10):
        sheet.cell(4, column).value = "style row"
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    workbook.close()


def _install_v3_database_settings(monkeypatch: pytest.MonkeyPatch) -> None:
    for key, value in REPORTING_ENV.items():
        monkeypatch.setenv(key, value)


@pytest.fixture
def v3_test_dirs():
    created: list[Path] = []

    def register(path: Path) -> Path:
        created.append(path)
        return path

    yield register
    for path in reversed(created):
        shutil.rmtree(path, ignore_errors=True)


def _sample_rows(sql_path: Path) -> list[dict[str, object]]:
    if sql_path.name == "sessions.sql":
        return [
            {"project_code": 1016, "username": "alice", "created_at": None, "closed_at": None,
             "count_retrieve_mode": 2, "count_pdf_viewed": 3, "count_content_copied": 4,
             "count_ask_mode": 5, "count_submission_times": 2, "total_issue_count": 6,
             "pending_issue": "1 / 6"},
            {"project_code": 1007, "username": "bob", "created_at": None, "closed_at": None,
             "count_retrieve_mode": 7, "count_pdf_viewed": 8, "count_content_copied": 9,
             "count_ask_mode": 10, "count_submission_times": 3, "total_issue_count": 11,
             "pending_issue": "2 / 11"},
            {"project_code": 0, "username": "ignored", "created_at": None, "closed_at": None,
             "count_retrieve_mode": 100, "count_pdf_viewed": 100, "count_content_copied": 100,
             "count_ask_mode": 100, "count_submission_times": 100, "total_issue_count": 100,
             "pending_issue": "0 / 0"},
        ]
    return [
        {"project_code": 1016, "count_submission_times": 12},
        {"project_code": 1007, "count_submission_times": 13},
        {"project_code": 2020, "count_submission_times": 4},
    ]


def test_weekly_report_is_independent_of_ecm_gates(tmp_path: Path) -> None:
    history_path = tmp_path / "history.sqlite3"
    from control_panel_v3.services.history import HistoryStore

    history = HistoryStore(history_path, tmp_path / "runs")
    history.initialize()
    period = history.create_report_period("2026-07-01", "2026-07-12")
    summary = history.list_report_period_summaries(REPORT_WORKFLOW_IDS)[0]
    assert ecm_workflow_gate_error(summary, "weekly-report") is None
    run = history.create_run("weekly-report", "Weekly report", period.start_date, period.end_date, period_id=period.id)
    assert run.workflow_id == "weekly-report"


def test_weekly_builder_uses_v3_template_run_directory_and_secret_free_command(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, v3_test_dirs
) -> None:
    _install_v3_database_settings(monkeypatch)
    settings = load_settings()
    root = v3_test_dirs(CONTROL_PANEL_ROOT / "tests" / "weekly-builder-fixture")
    template = root / "templates" / "weekly.xlsx"
    _make_template(template)
    test_settings = type(settings)(**{
        **settings.__dict__,
        "data_dir": root / "data",
        "archive_dir": root / "archive",
        "csv_dir": root / "csv",
        "ecm_output_dir": root / "ecm",
        "smm_output_dir": root / "smm",
        "django_output_dir": root / "django",
        "edge_profile_dir": root / "profile",
        "weekly_template_path": template,
    })
    prepared = prepare_workflow("weekly-report", date(2026, 7, 1), date(2026, 7, 12), 902, test_settings)
    command = " ".join(prepared.command)
    assert "control_panel_v3.workers.weekly_report" in command
    assert str(template) in command
    assert "weekly-run-902" in command
    assert REPORTING_ENV["CONTROL_PANEL_V3_WEEKLY_PGPASSWORD"] not in command
    assert prepared.artifacts[0].path.is_relative_to(root / "django" / "weekly-run-902")
    assert prepared.environment_overrides is not None
    assert prepared.environment_overrides["PGPASSWORD"] == ""
    assert prepared.environment_overrides["CONTROL_PANEL_V3_WEEKLY_PGPASSWORD"] == REPORTING_ENV["CONTROL_PANEL_V3_WEEKLY_PGPASSWORD"]


def test_weekly_builder_requires_v3_credentials_only_for_weekly(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, v3_test_dirs
) -> None:
    for name in REPORTING_ENV:
        monkeypatch.delenv(name, raising=False)
    settings = load_settings()
    fixture_root = v3_test_dirs(CONTROL_PANEL_ROOT / "tests" / "weekly-missing-config")
    weekly_template = fixture_root / "weekly.xlsx"
    _make_template(weekly_template)
    test_settings = type(settings)(**{
        **settings.__dict__,
        "weekly_template_path": weekly_template,
        "django_output_dir": CONTROL_PANEL_ROOT / "tests" / "weekly-missing-config" / "out",
    })
    with pytest.raises(PreflightError, match="V3 weekly database settings"):
        prepare_workflow("weekly-report", date(2026, 7, 1), date(2026, 7, 12), 903, test_settings)
    assert settings.validate() == []


def test_weekly_worker_creates_atomic_report_and_never_uses_ambient_database_settings(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, v3_test_dirs
) -> None:
    from control_panel_v3.workers.weekly_report import _load_reference_report

    template_dir = v3_test_dirs(CONTROL_PANEL_ROOT / "tests" / "weekly-worker-template")
    template = template_dir / "template.xlsx"
    _make_template(template)
    template_hash = _sha256(template)
    output_dir = v3_test_dirs(CONTROL_PANEL_ROOT / "tests" / "weekly-worker-output")
    _install_v3_database_settings(monkeypatch)
    monkeypatch.setenv("PGHOST", "ambient-host-must-not-win")
    monkeypatch.setenv("PGDATABASE", "ambient-database-must-not-win")
    monkeypatch.setenv("PGUSER", "ambient-user-must-not-win")
    monkeypatch.setenv("PGPASSWORD", "ambient-password-must-not-win")
    monkeypatch.setenv("REPORTING_TIMEZONE", "UTC")
    calls: list[tuple[Path, tuple[date, date], str | None, str | None, str | None, str | None, str | None]] = []
    report = _load_reference_report()

    def fake_fetch_rows(sql_path: Path, parameters: tuple[date, date]) -> list[dict[str, object]]:
        calls.append((
            sql_path,
            parameters,
            os.getenv("PGHOST"),
            os.getenv("PGDATABASE"),
            os.getenv("PGUSER"),
            os.getenv("PGPASSWORD"),
            os.getenv("REPORTING_TIMEZONE"),
        ))
        return _sample_rows(sql_path)

    report.fetch_rows = fake_fetch_rows
    report.load_dotenv = lambda *_args, **_kwargs: pytest.fail("reference dotenv loader was called")
    output = generate_weekly_report(
        date(2026, 7, 1), date(2026, 7, 12), template, output_dir,
        reference_report=report,
    )
    assert output.is_file()
    assert output.parent == output_dir
    assert len(calls) == 2
    assert all(call[1] == (date(2026, 7, 1), date(2026, 7, 12)) for call in calls)
    assert all(call[2:] == (
        REPORTING_ENV["CONTROL_PANEL_V3_WEEKLY_PGHOST"],
        REPORTING_ENV["CONTROL_PANEL_V3_WEEKLY_PGDATABASE"],
        REPORTING_ENV["CONTROL_PANEL_V3_WEEKLY_PGUSER"],
        REPORTING_ENV["CONTROL_PANEL_V3_WEEKLY_PGPASSWORD"],
        "Asia/Hong_Kong",
    ) for call in calls)
    assert _sha256(template) == template_hash

    workbook = load_workbook(output, data_only=False)
    try:
        assert workbook.sheetnames == ["1 - 12 Jul - Consolidated", "Sessions", "Submission Counts"]
        summary = workbook["1 - 12 Jul - Consolidated"]
        assert summary["A1"].value == "Template title"
        assert summary["A2"].value == "Template subtitle"
        assert summary.cell(4, 1).value == "1016 (aka 1007)"
        assert summary.cell(4, 2).value == 2
        assert summary.cell(4, 6).value == 15
        assert summary.cell(4, 7).value == 25
        assert summary.cell(4, 8).value == 17
        assert summary.cell(5, 1).value == 2020
        assert summary.cell(5, 2).value == 0
        assert summary.cell(5, 6).value == 0
        assert summary.cell(5, 7).value == 4
        assert workbook["Sessions"].max_row == 3
        assert workbook["Sessions"].cell(1, 1).value == "project_code"
        assert workbook["Submission Counts"].max_row == 4
    finally:
        workbook.close()

    with pytest.raises(WeeklyReportError, match="refusing to overwrite"):
        generate_weekly_report(
            date(2026, 7, 1), date(2026, 7, 12), template, output_dir,
            reference_report=report,
        )


def test_weekly_worker_rejects_missing_v3_credentials_even_if_generic_values_exist(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, v3_test_dirs
) -> None:
    from control_panel_v3.workers.weekly_report import _load_reference_report

    for name in REPORTING_ENV:
        monkeypatch.delenv(name, raising=False)
    monkeypatch.setenv("PGHOST", "ambient-host")
    monkeypatch.setenv("PGDATABASE", "ambient-database")
    monkeypatch.setenv("PGUSER", "ambient-user")
    monkeypatch.setenv("PGPASSWORD", "ambient-password")
    template_dir = v3_test_dirs(CONTROL_PANEL_ROOT / "tests" / "weekly-missing-credentials-template")
    template = template_dir / "template.xlsx"
    _make_template(template)
    output_dir = v3_test_dirs(CONTROL_PANEL_ROOT / "tests" / "weekly-missing-credentials-output")
    report = _load_reference_report()
    with pytest.raises(WeeklyReportError, match="CONTROL_PANEL_V3_WEEKLY_PGHOST"):
        generate_weekly_report(
            date(2026, 7, 1), date(2026, 7, 12), template, output_dir,
            reference_report=report,
        )
    assert not output_dir.exists()
