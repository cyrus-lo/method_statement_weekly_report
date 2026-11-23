from __future__ import annotations

import hashlib
import os
import shutil
from datetime import date
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from control_panel_v3.services import config
from control_panel_v3.services.config import (
    CONTROL_PANEL_ROOT,
    ConfigurationError,
    load_settings,
    load_weekly_report_environment,
)
from control_panel_v3.services.workflow_builders import PreflightError, prepare_workflow
from control_panel_v3.services.workflows import REPORT_WORKFLOW_IDS, ecm_workflow_gate_error
from control_panel_v3.workers.weekly_report import WeeklyReportError, generate_weekly_report


REPORTING_ENV = {
    "PGHOST": "django-db.example.test",
    "PGDATABASE": "weekly_test",
    "PGUSER": "read_only_test",
    "PGPASSWORD": "test-password-not-a-real-secret",
    "PGPORT": "5433",
    "PGSSLMODE": "require",
    "REPORTING_TIMEZONE": "Asia/Hong_Kong",
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


def _write_weekly_env(path: Path, values: dict[str, str] | None = None) -> Path:
    settings = values or REPORTING_ENV
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("".join(f"{key}={value}\n" for key, value in settings.items()), encoding="utf-8")
    return path


def _use_weekly_env(monkeypatch: pytest.MonkeyPatch, path: Path) -> None:
    monkeypatch.setattr(config, "DJANGO_ENV_PATH", path)


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


def test_weekly_env_parser_uses_file_values_without_mutating_process_environment(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    env_file = _write_weekly_env(tmp_path / "django.env")
    for key in REPORTING_ENV:
        monkeypatch.setenv(key, "ambient-must-not-win")
    before = {key: os.environ.get(key) for key in REPORTING_ENV}
    result = load_weekly_report_environment(env_file)
    assert result == REPORTING_ENV
    assert {key: os.environ.get(key) for key in REPORTING_ENV} == before


@pytest.mark.parametrize("missing_key", ("PGHOST", "PGDATABASE", "PGUSER", "PGPASSWORD"))
def test_weekly_env_parser_rejects_missing_or_blank_required_settings(
    tmp_path: Path, missing_key: str
) -> None:
    values = dict(REPORTING_ENV)
    values[missing_key] = "   "
    with pytest.raises(ConfigurationError, match=missing_key):
        load_weekly_report_environment(_write_weekly_env(tmp_path / "django.env", values))


@pytest.mark.parametrize("port", ("0", "65536", "not-a-port"))
def test_weekly_env_parser_rejects_invalid_port(tmp_path: Path, port: str) -> None:
    values = dict(REPORTING_ENV, PGPORT=port)
    with pytest.raises(ConfigurationError, match="PGPORT"):
        load_weekly_report_environment(_write_weekly_env(tmp_path / "django.env", values))


def test_weekly_env_parser_applies_optional_defaults_and_validates_timezone(tmp_path: Path) -> None:
    values = dict(REPORTING_ENV)
    values.pop("PGPORT")
    values.pop("PGSSLMODE")
    values.pop("REPORTING_TIMEZONE")
    parsed = load_weekly_report_environment(_write_weekly_env(tmp_path / "defaults.env", values))
    assert parsed["PGPORT"] == "5432"
    assert "PGSSLMODE" not in parsed
    assert parsed["REPORTING_TIMEZONE"] == "Asia/Hong_Kong"

    values["REPORTING_TIMEZONE"] = "   "
    with pytest.raises(ConfigurationError, match="REPORTING_TIMEZONE"):
        load_weekly_report_environment(_write_weekly_env(tmp_path / "blank-timezone.env", values))

    values["REPORTING_TIMEZONE"] = "not/a-zone"
    with pytest.raises(ConfigurationError, match="REPORTING_TIMEZONE"):
        load_weekly_report_environment(_write_weekly_env(tmp_path / "invalid-timezone.env", values))


def test_weekly_builder_uses_v3_template_run_directory_and_secret_free_command(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, v3_test_dirs
) -> None:
    env_file = _write_weekly_env(tmp_path / "django.env")
    _use_weekly_env(monkeypatch, env_file)
    for key in REPORTING_ENV:
        monkeypatch.setenv(key, "ambient-must-not-win")
    for key in REPORTING_ENV:
        monkeypatch.setenv(f"CONTROL_PANEL_V3_WEEKLY_{key}", "obsolete-must-not-win")
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
    assert REPORTING_ENV["PGPASSWORD"] not in command
    assert prepared.artifacts[0].path.is_relative_to(root / "django" / "weekly-run-902")
    assert prepared.environment_overrides is not None
    assert prepared.environment_overrides["PGPASSWORD"] == REPORTING_ENV["PGPASSWORD"]
    assert prepared.environment_overrides["PGHOST"] == REPORTING_ENV["PGHOST"]
    assert prepared.environment_overrides["PGPORT"] == "5433"
    assert prepared.environment_overrides["REPORTING_TIMEZONE"] == REPORTING_ENV["REPORTING_TIMEZONE"]
    assert "CONTROL_PANEL_V3_WEEKLY_PGPASSWORD" not in prepared.environment_overrides
    assert set(prepared.environment_overrides) == {
        "PATH", "SYSTEMROOT", "WINDIR", "TEMP", "TMP", *REPORTING_ENV
    }


def test_weekly_builder_requires_django_env_only_for_weekly(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, v3_test_dirs
) -> None:
    _use_weekly_env(monkeypatch, tmp_path / "missing.env")
    settings = load_settings()
    fixture_root = v3_test_dirs(CONTROL_PANEL_ROOT / "tests" / "weekly-missing-config")
    weekly_template = fixture_root / "weekly.xlsx"
    _make_template(weekly_template)
    test_settings = type(settings)(**{
        **settings.__dict__,
        "weekly_template_path": weekly_template,
        "django_output_dir": CONTROL_PANEL_ROOT / "tests" / "weekly-missing-config" / "out",
    })
    with pytest.raises(PreflightError, match="django/.env setting") as error:
        prepare_workflow("weekly-report", date(2026, 7, 1), date(2026, 7, 12), 903, test_settings)
    assert REPORTING_ENV["PGPASSWORD"] not in str(error.value)
    assert settings.validate() == []


def test_weekly_worker_creates_atomic_report_from_isolated_child_settings(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, v3_test_dirs, capsys: pytest.CaptureFixture[str]
) -> None:
    from control_panel_v3.workers.weekly_report import _load_reference_report

    template_dir = v3_test_dirs(CONTROL_PANEL_ROOT / "tests" / "weekly-worker-template")
    template = template_dir / "template.xlsx"
    _make_template(template)
    template_hash = _sha256(template)
    output_dir = v3_test_dirs(CONTROL_PANEL_ROOT / "tests" / "weekly-worker-output")
    for key, value in REPORTING_ENV.items():
        monkeypatch.setenv(key, value)
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
        REPORTING_ENV["PGHOST"],
        REPORTING_ENV["PGDATABASE"],
        REPORTING_ENV["PGUSER"],
        REPORTING_ENV["PGPASSWORD"],
        REPORTING_ENV["REPORTING_TIMEZONE"],
    ) for call in calls)
    assert {key: os.getenv(key) for key in REPORTING_ENV} == REPORTING_ENV
    assert REPORTING_ENV["PGPASSWORD"] not in capsys.readouterr().out
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


def test_weekly_worker_rejects_missing_generic_credentials_even_if_obsolete_values_exist(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, v3_test_dirs
) -> None:
    from control_panel_v3.workers.weekly_report import _load_reference_report

    for name in (*REPORTING_ENV, "CONTROL_PANEL_V3_WEEKLY_PGHOST", "CONTROL_PANEL_V3_WEEKLY_PGDATABASE",
                 "CONTROL_PANEL_V3_WEEKLY_PGUSER", "CONTROL_PANEL_V3_WEEKLY_PGPASSWORD"):
        monkeypatch.delenv(name, raising=False)
    monkeypatch.setenv("CONTROL_PANEL_V3_WEEKLY_PGHOST", "obsolete-host")
    monkeypatch.setenv("CONTROL_PANEL_V3_WEEKLY_PGDATABASE", "obsolete-database")
    monkeypatch.setenv("CONTROL_PANEL_V3_WEEKLY_PGUSER", "obsolete-user")
    monkeypatch.setenv("CONTROL_PANEL_V3_WEEKLY_PGPASSWORD", "obsolete-password")
    template_dir = v3_test_dirs(CONTROL_PANEL_ROOT / "tests" / "weekly-missing-credentials-template")
    template = template_dir / "template.xlsx"
    _make_template(template)
    output_dir = v3_test_dirs(CONTROL_PANEL_ROOT / "tests" / "weekly-missing-credentials-output")
    report = _load_reference_report()
    with pytest.raises(WeeklyReportError, match="PGHOST"):
        generate_weekly_report(
            date(2026, 7, 1), date(2026, 7, 12), template, output_dir,
            reference_report=report,
        )
    assert not output_dir.exists()
