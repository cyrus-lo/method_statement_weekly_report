from __future__ import annotations

import sqlite3
from datetime import date
from pathlib import Path

import pytest

from control_panel.services.history import HistoryStore
from control_panel.services.periods import latest_completed_report_period, report_period_for

from .test_workflows import build_settings


WORKFLOW_IDS = ("weekly-report", "ecm-step-1a", "ecm-step-1b", "ecm-step-2")


def test_report_periods_split_and_join_at_month_end() -> None:
    assert report_period_for(date(2026, 6, 22)) == (date(2026, 6, 22), date(2026, 6, 30))
    assert report_period_for(date(2026, 6, 30)) == (date(2026, 6, 22), date(2026, 6, 30))
    assert report_period_for(date(2026, 7, 1)) == (date(2026, 7, 1), date(2026, 7, 12))
    assert report_period_for(date(2026, 7, 12)) == (date(2026, 7, 1), date(2026, 7, 12))
    assert report_period_for(date(2026, 7, 13)) == (date(2026, 7, 13), date(2026, 7, 19))


def test_report_periods_handle_month_start_on_monday_sunday_and_year_end() -> None:
    assert report_period_for(date(2026, 6, 1)) == (date(2026, 6, 1), date(2026, 6, 7))
    assert report_period_for(date(2026, 2, 1)) == (date(2026, 2, 1), date(2026, 2, 8))
    assert report_period_for(date(2026, 12, 28)) == (date(2026, 12, 21), date(2026, 12, 31))
    assert report_period_for(date(2027, 1, 1)) == (date(2027, 1, 1), date(2027, 1, 10))


def test_latest_completed_period_does_not_select_an_active_month_end_period() -> None:
    assert latest_completed_report_period(date(2026, 7, 27)) == (
        date(2026, 7, 13),
        date(2026, 7, 19),
    )


def test_report_period_summary_uses_latest_run_per_workflow_and_locks_dates(tmp_path: Path) -> None:
    settings = build_settings(tmp_path)
    history = HistoryStore(settings.database_path, settings.runs_dir)
    history.initialize()
    period = history.create_report_period("2026-07-01", "2026-07-12")

    first_weekly = history.create_run(
        "weekly-report", "Weekly report", period.start_date, period.end_date, period_id=period.id
    )
    history.mark_finished(first_weekly.id, state="failed", exit_code=1, error_summary="Test failure")
    latest_weekly = history.create_run(
        "weekly-report", "Weekly report", period.start_date, period.end_date, period_id=period.id
    )
    history.mark_finished(latest_weekly.id, state="succeeded", exit_code=0)
    step_1a = history.create_run(
        "ecm-step-1a", "ECM Step 1a", period.start_date, period.end_date, period_id=period.id
    )

    ungrouped = history.create_run(
        "ecm-step-1b", "Flexible dashboard run", "2026-07-20", "2026-07-26"
    )
    summaries = history.list_report_period_summaries(WORKFLOW_IDS)

    assert len(summaries) == 1
    summary = summaries[0]
    assert summary.is_locked
    assert summary.runs["weekly-report"] == history.get_run(latest_weekly.id)
    assert summary.runs["ecm-step-1a"] == history.get_run(step_1a.id)
    assert summary.runs["ecm-step-1b"] is None
    assert history.get_run(ungrouped.id).period_id is None
    assert summary.last_updated_at.endswith("+08:00")

    with pytest.raises(ValueError, match="locked"):
        history.update_report_period(period.id, "2026-07-02", "2026-07-12")


def test_report_period_can_be_edited_before_its_first_workflow(tmp_path: Path) -> None:
    settings = build_settings(tmp_path)
    history = HistoryStore(settings.database_path, settings.runs_dir)
    history.initialize()
    period = history.create_report_period("2026-07-13", "2026-07-19")

    updated = history.update_report_period(period.id, "2026-07-14", "2026-07-20")

    assert (updated.start_date, updated.end_date) == ("2026-07-14", "2026-07-20")


def test_report_period_can_be_soft_deleted_restored_and_recreated(tmp_path: Path) -> None:
    settings = build_settings(tmp_path)
    history = HistoryStore(settings.database_path, settings.runs_dir)
    history.initialize()
    period = history.create_report_period("2026-07-13", "2026-07-19")
    run = history.create_run(
        "weekly-report", "Weekly report", period.start_date, period.end_date, period_id=period.id
    )
    history.mark_finished(run.id, state="succeeded", exit_code=0)

    deleted = history.delete_report_period(period.id)

    assert deleted.is_deleted
    assert history.get_run(run.id).state == "succeeded"
    assert history.list_report_period_summaries(WORKFLOW_IDS) == []
    assert [summary.period.id for summary in history.list_report_period_summaries(
        WORKFLOW_IDS, is_deleted=True
    )] == [period.id]

    replacement = history.create_report_period("2026-07-13", "2026-07-19")
    with pytest.raises(ValueError, match="active report period already uses these dates"):
        history.restore_report_period(period.id)

    assert [summary.period.id for summary in history.list_report_period_summaries(WORKFLOW_IDS)] == [
        replacement.id,
    ]


def test_report_period_cannot_be_deleted_while_a_workflow_is_active(tmp_path: Path) -> None:
    settings = build_settings(tmp_path)
    history = HistoryStore(settings.database_path, settings.runs_dir)
    history.initialize()
    period = history.create_report_period("2026-07-13", "2026-07-19")
    history.create_run(
        "weekly-report", "Weekly report", period.start_date, period.end_date, period_id=period.id
    )

    with pytest.raises(ValueError, match="cannot be deleted while a workflow is running"):
        history.delete_report_period(period.id)

    assert not history.get_report_period(period.id).is_deleted


def test_initialize_migrates_an_existing_runs_table_without_losing_runs(tmp_path: Path) -> None:
    database_path = tmp_path / "data" / "history.sqlite3"
    database_path.parent.mkdir()
    with sqlite3.connect(database_path) as connection:
        connection.executescript(
            """
            CREATE TABLE runs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                workflow_id TEXT NOT NULL,
                workflow_label TEXT NOT NULL,
                start_date TEXT NOT NULL,
                end_date TEXT NOT NULL,
                state TEXT NOT NULL,
                created_at TEXT NOT NULL,
                started_at TEXT,
                finished_at TEXT,
                exit_code INTEGER,
                error_summary TEXT,
                log_path TEXT NOT NULL
            );
            INSERT INTO runs (
                workflow_id, workflow_label, start_date, end_date, state, created_at, log_path
            ) VALUES (
                'weekly-report', 'Weekly report', '2026-07-01', '2026-07-12',
                'succeeded', '2026-07-01T01:00:00+00:00', 'run.log'
            );
            """
        )

    history = HistoryStore(database_path, tmp_path / "data" / "runs")
    history.initialize()

    migrated = history.get_run(1)
    assert migrated.period_id is None
    assert migrated.created_at == "2026-07-01T09:00:00+08:00"
    assert history.create_report_period("2026-07-01", "2026-07-12").id == 1


def test_initialize_migrates_legacy_report_periods_without_losing_related_records(
    tmp_path: Path,
) -> None:
    database_path = tmp_path / "data" / "history.sqlite3"
    database_path.parent.mkdir()
    with sqlite3.connect(database_path) as connection:
        connection.executescript(
            """
            CREATE TABLE report_periods (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                start_date TEXT NOT NULL,
                end_date TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(start_date, end_date)
            );
            CREATE TABLE runs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                period_id INTEGER REFERENCES report_periods(id),
                workflow_id TEXT NOT NULL,
                workflow_label TEXT NOT NULL,
                start_date TEXT NOT NULL,
                end_date TEXT NOT NULL,
                state TEXT NOT NULL,
                created_at TEXT NOT NULL,
                started_at TEXT,
                finished_at TEXT,
                exit_code INTEGER,
                error_summary TEXT,
                log_path TEXT NOT NULL
            );
            CREATE TABLE artifacts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                run_id INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
                name TEXT NOT NULL,
                source_path TEXT NOT NULL,
                stored_path TEXT NOT NULL,
                size_bytes INTEGER NOT NULL
            );
            INSERT INTO report_periods VALUES (1, '2026-07-13', '2026-07-19',
                '2026-07-20T00:00:00+08:00', '2026-07-20T00:00:00+08:00');
            INSERT INTO runs VALUES (1, 1, 'weekly-report', 'Weekly report',
                '2026-07-13', '2026-07-19', 'succeeded', '2026-07-20T00:00:00+08:00',
                '2026-07-20T00:01:00+08:00', '2026-07-20T00:02:00+08:00', 0, NULL, 'run.log');
            INSERT INTO artifacts VALUES (1, 1, 'report.xlsx', 'source.xlsx', 'stored.xlsx', 10);
            """
        )

    history = HistoryStore(database_path, tmp_path / "data" / "runs")
    history.initialize()

    period = history.get_report_period(1)
    assert not period.is_deleted
    assert history.get_run(1).period_id == period.id
    assert history.get_artifact(1, 1).name == "report.xlsx"
    assert history.delete_report_period(period.id).is_deleted
    assert history.create_report_period(period.start_date, period.end_date).id == 2

    with sqlite3.connect(database_path) as connection:
        foreign_key = connection.execute("PRAGMA foreign_key_list(runs)").fetchone()
    assert foreign_key[2] == "report_periods"
