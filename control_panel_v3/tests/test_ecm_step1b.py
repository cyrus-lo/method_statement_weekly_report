from __future__ import annotations

import shutil
from datetime import date, datetime
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from control_panel_v3.services.config import CONTROL_PANEL_ROOT, load_settings
from control_panel_v3.services.workflow_builders import PreflightError, prepare_workflow
from control_panel_v3.services.history import HistoryStore
from control_panel_v3.services.workflows import REPORT_WORKFLOW_IDS, ecm_workflow_gate_error
from control_panel_v3.workers.ecm_step1a import OUTGOING_REFERENCE_HEADER, REVIEW_HEADERS
from control_panel_v3.workers.ecm_step1b import Step1bError, run_step1b


SHEET = "ECM_1-3 May (WIP)"


def _workbook(path: Path, values: tuple[tuple[object, ...], ...]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = SHEET
    for column, header in enumerate(REVIEW_HEADERS, 1):
        sheet.cell(1, column).value = header
    sheet.cell(1, 13).value = "Remark"
    sheet.cell(1, 16).value = OUTGOING_REFERENCE_HEADER
    for row_number, row in enumerate(values, 2):
        for column, value in enumerate(row, 1):
            sheet.cell(row_number, column).value = value
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    workbook.close()


def _row(name: str, count: str | None, revision: str | None, remark: object) -> tuple[object, ...]:
    return (1016, name, date(2026, 5, 1), f"CWCMGCJV/1016/CSF/{name}.pdf", revision, "Operator", 1, "SMM", count, "Y", "N", "keep L", remark)


def _under_v3(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    root = tmp_path / "control_panel_v3"
    root.mkdir()
    monkeypatch.setattr("control_panel_v3.workers.ecm_step1b.CONTROL_PANEL_ROOT", root)
    return root


def test_step1b_strict_cutoff_and_snapshots_are_transactional(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    root = _under_v3(tmp_path, monkeypatch)
    workbook_path = root / "working" / "ecm" / "ECM_Record.xlsx"
    backup = root / "working" / "ecm" / "backups" / "before.xlsx"
    post = root / "working" / "ecm" / "snapshots" / "after.xlsx"
    rows = (
        _row("After", None, "A", "2026-04-17"),
        _row("Cutoff", None, "B", datetime(2026, 4, 16)),
        _row("No revision", None, "  ", "keep this remark"),
        _row("Already N", "N", "C", None),
    )
    _workbook(workbook_path, rows)
    original_bytes = workbook_path.read_bytes()
    calls: list[tuple[Path, str]] = []

    def fake_excel_filter(staged: Path, worksheet: str) -> None:
        calls.append((staged, worksheet))

    counts = run_step1b(workbook_path, SHEET, backup, post, excel_filter=fake_excel_filter)
    assert counts == (2, 1)
    assert calls and calls[0][0] != workbook_path and calls[0][1] == SHEET
    assert backup.read_bytes() == original_bytes
    assert post.read_bytes() == workbook_path.read_bytes()

    workbook = load_workbook(workbook_path, data_only=False)
    try:
        sheet = workbook[SHEET]
        assert [sheet.cell(row, 9).value for row in range(2, 6)] == ["Y", "N", "Y", "N"]
        assert sheet.cell(2, 13).value == "1st version is submitted on 17 Apr 2026"
        assert sheet.cell(3, 13).value == "1st version is submitted on 16 Apr 2026"
        assert sheet.cell(4, 13).value == "keep this remark"
        assert sheet.cell(5, 13).value is None
        assert sheet.auto_filter.ref == "A1:P5"
        assert sheet.auto_filter.filterColumn[0].filters.filter == ["Y"]
    finally:
        workbook.close()


def test_step1b_validates_all_rows_before_any_durable_change(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    root = _under_v3(tmp_path, monkeypatch)
    workbook_path = root / "ECM_Record.xlsx"
    backup = root / "before.xlsx"
    post = root / "after.xlsx"
    _workbook(workbook_path, (_row("valid first", None, "A", "2026-04-17"), _row("missing date", None, "B", None)))
    original = workbook_path.read_bytes()
    called = False

    def should_not_run(_path: Path, _sheet: str) -> None:
        nonlocal called
        called = True

    with pytest.raises(Step1bError, match="requires a verified"):
        run_step1b(workbook_path, SHEET, backup, post, excel_filter=should_not_run)
    assert workbook_path.read_bytes() == original
    assert not backup.exists() and not post.exists()
    assert not called


def test_step1b_excel_failure_keeps_active_workbook_unchanged_and_no_post_snapshot(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    root = _under_v3(tmp_path, monkeypatch)
    workbook_path = root / "ECM_Record.xlsx"
    backup = root / "backups" / "before.xlsx"
    post = root / "snapshots" / "after.xlsx"
    _workbook(workbook_path, (_row("candidate", None, "A", "2026-04-17"),))
    original = workbook_path.read_bytes()

    def fail_excel(_path: Path, _sheet: str) -> None:
        raise RuntimeError("Excel filter failed")

    with pytest.raises(RuntimeError, match="Excel filter failed"):
        run_step1b(workbook_path, SHEET, backup, post, excel_filter=fail_excel)
    assert workbook_path.read_bytes() == original
    assert backup.read_bytes() == original
    assert not post.exists()


def test_step1b_rejects_existing_evidence_and_paths_outside_v3(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    root = _under_v3(tmp_path, monkeypatch)
    workbook_path = root / "ECM_Record.xlsx"
    backup = root / "backup.xlsx"
    post = root / "post.xlsx"
    _workbook(workbook_path, (_row("candidate", None, "A", "2026-04-17"),))
    with pytest.raises(Step1bError, match="must remain under"):
        run_step1b(tmp_path / "outside.xlsx", SHEET, backup, post, excel_filter=lambda *_: None)
    backup.write_text("protected evidence", encoding="utf-8")
    with pytest.raises(Step1bError, match="refusing to overwrite"):
        run_step1b(workbook_path, SHEET, backup, post, excel_filter=lambda *_: None)


def test_step1b_builder_requires_v3_workbook_and_collects_all_snapshots(tmp_path: Path) -> None:
    settings = load_settings()
    roots = {name: settings.project_root / "control_panel_v3" / "tests" / f"step1b-{name}" for name in (
        "data_dir", "archive_dir", "csv_dir", "ecm_output_dir", "smm_output_dir", "django_output_dir", "edge_profile_dir"
    )}
    test_settings = type(settings)(**{**settings.__dict__, **roots})
    test_settings.ensure_application_directories()
    workbook_path = test_settings.ecm_output_dir / "ECM_Record.xlsx"
    _workbook(workbook_path, (_row("candidate", None, "A", "2026-04-17"),))
    prepared = prepare_workflow("ecm-step-1b", date(2026, 5, 1), date(2026, 5, 3), 81, test_settings)
    assert "control_panel_v3.workers.ecm_step1b" in prepared.command
    assert str(workbook_path) in prepared.command
    assert "--post-backup" in prepared.command
    assert "--excel-filter-timeout-seconds" in prepared.command
    assert len(prepared.artifacts) == 3
    assert all(source.path.is_relative_to(CONTROL_PANEL_ROOT) for source in prepared.artifacts)
    (test_settings.ecm_output_dir / "ECM_Record.xlsx").unlink()
    with pytest.raises(PreflightError, match="ECM review workbook"):
        prepare_workflow("ecm-step-1b", date(2026, 5, 1), date(2026, 5, 3), 82, test_settings)
    for directory in roots.values():
        shutil.rmtree(directory, ignore_errors=True)


def test_step1b_gate_requires_a_successful_lookup(tmp_path: Path) -> None:
    history = HistoryStore(tmp_path / "history.sqlite3", tmp_path / "runs")
    history.initialize()
    period = history.create_report_period("2026-05-01", "2026-05-03")
    step0 = history.create_run("ecm-step-0", "Step 0", period.start_date, period.end_date, period_id=period.id)
    history.mark_finished(step0.id, state="succeeded", exit_code=0)
    step1a = history.create_run("ecm-step-1a", "Step 1a", period.start_date, period.end_date, period_id=period.id)
    history.mark_finished(step1a.id, state="succeeded", exit_code=0)
    summary = history.list_report_period_summaries(REPORT_WORKFLOW_IDS)[0]
    assert "lookup must succeed" in (ecm_workflow_gate_error(summary, "ecm-step-1b") or "")
    lookup = history.create_run("ecm-first-date-lookup", "Lookup", period.start_date, period.end_date, period_id=period.id)
    history.mark_finished(lookup.id, state="succeeded", exit_code=0)
    summary = history.list_report_period_summaries(REPORT_WORKFLOW_IDS)[0]
    assert ecm_workflow_gate_error(summary, "ecm-step-1b") is None
    failed_one_b = history.create_run("ecm-step-1b", "Step 1b", period.start_date, period.end_date, period_id=period.id)
    history.mark_finished(failed_one_b.id, state="failed", exit_code=1, error_summary="Filter failed")
    summary = history.list_report_period_summaries(REPORT_WORKFLOW_IDS)[0]
    assert ecm_workflow_gate_error(summary, "ecm-step-2", confirm_step_1b_bypass=True) is not None