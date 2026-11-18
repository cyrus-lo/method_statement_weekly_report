from __future__ import annotations

import hashlib
import shutil
from datetime import date
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

from control_panel_v3.services.config import CONTROL_PANEL_ROOT, load_settings
from control_panel_v3.services.history import HistoryStore
from control_panel_v3.services.workflow_builders import PreflightError, prepare_workflow
from control_panel_v3.services.workflows import REPORT_WORKFLOW_IDS, ecm_workflow_gate_error
from control_panel_v3.workers.ecm_step1b import OUTGOING_REFERENCE_HEADER, REVIEW_HEADERS
from control_panel_v3.workers.ecm_step2 import (
    READINESS_SHEET_NAME,
    Step2Error,
    aggregate_ecm_record,
    run_step2,
    summary_sheet_name,
)


ECM_SHEET = "ECM_1-13 Jul (WIP)"
LAYOUT_SHEET = "(WIP) 1 Jul - 12 Jul"
SUMMARY_SHEET = "(WIP) 1 Jul - 13 Jul 2026"


def _ecm_row(project: object, count: object = "Y", use_smm: object = "Y", use_kms: object = "Y", remark: object = None) -> tuple[object, ...]:
    return (project, "Document", date(2026, 7, 1), "Document.pdf", "A", "Operator", 1, "SMM-1", count, use_smm, use_kms, None, remark, None)


def _make_ecm(path: Path, rows: tuple[tuple[object, ...], ...], *, extra_sheet: bool = True) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = ECM_SHEET
    for column, header in enumerate(REVIEW_HEADERS, 1):
        sheet.cell(1, column).value = header
    sheet.cell(1, 12).value = "Work Flow"
    sheet.cell(1, 13).value = "Remark"
    sheet.cell(1, 14).value = "Display"
    sheet.cell(1, 16).value = OUTGOING_REFERENCE_HEADER
    for row_number, row in enumerate(rows, 2):
        for column, value in enumerate(row, 1):
            sheet.cell(row_number, column).value = value
    sheet.auto_filter.ref = f"A1:P{sheet.max_row}"
    if len(rows) > 1:
        sheet.row_dimensions[3].hidden = True
    if extra_sheet:
        workbook.create_sheet("Unrelated") ["A1"] = "preserve me"
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    workbook.close()


def _layout_sheet(workbook: Workbook, title: str = LAYOUT_SHEET) -> None:
    sheet = workbook.create_sheet(title)
    sheet.cell(2, 2).value = "Project"
    sheet.cell(2, 3).value = "BU"
    sheet.cell(2, 4).value = "SMM (Y/N)"
    sheet.cell(2, 5).value = "Number of Method Statement Record in SMM"
    sheet.cell(2, 8).value = "Number of Method Statement Record (Outbound) in ECM"
    sheet.cell(3, 5).value = "A. Draft via KMS"
    sheet.cell(3, 6).value = "B. Draft NOT via KMS"
    sheet.cell(3, 7).value = "Total No. (A + B) Approved"
    for column in range(1, 9):
        sheet.cell(5, column).font = Font(color="FF000000")
        sheet.cell(6, column).font = Font(color="FF000000")
    sheet.merge_cells("E2:G2")
    sheet["A1"] = "layout preserved"


def _make_smm(path: Path, *, layout_title: str = LAYOUT_SHEET, readiness: tuple[tuple[object, str, str], ...] = ((101, "101 Project 101", "BU-101"), (1016, "1016 - 1/WSD/19 (=1007)", "BU-1016"), (200, "200 Project 200", "BU-200"))) -> None:
    workbook = Workbook()
    readiness_sheet = workbook.active
    readiness_sheet.title = READINESS_SHEET_NAME
    readiness_sheet.cell(1, 3).value = "Project"
    readiness_sheet.cell(1, 4).value = "BU"
    for row_number, (code, label, bu) in enumerate(readiness, 2):
        readiness_sheet.cell(row_number, 3).value = label or code
        readiness_sheet.cell(row_number, 4).value = bu
    _layout_sheet(workbook, layout_title)
    other = workbook.create_sheet("Unrelated")
    other["A1"] = "unchanged"
    other["B2"] = "=1+1"
    workbook.active = 0
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    workbook.close()


def _make_readiness(path: Path, readiness: tuple[tuple[object, str, str], ...] = ((101, "101 Project 101", "BU-101"), (1016, "1016 - 1/WSD/19 (=1007)", "BU-1016"), (200, "200 Project 200", "BU-200"))) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = READINESS_SHEET_NAME
    sheet.cell(1, 3).value = "Project"
    sheet.cell(1, 4).value = "BU"
    for row_number, (code, label, bu) in enumerate(readiness, 2):
        sheet.cell(row_number, 3).value = label or code
        sheet.cell(row_number, 4).value = bu
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    workbook.close()


def _v3_paths(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> tuple[Path, Path, Path, Path, Path]:
    root = tmp_path / "control_panel_v3"
    monkeypatch.setattr("control_panel_v3.workers.ecm_step2.CONTROL_PANEL_ROOT", root)
    ecm = root / "working" / "ecm" / "ECM_Record.xlsx"
    smm = root / "working" / "smm" / "SMM_KMS_PM_Master_Document.xlsx"
    template = root / "working" / "templates" / "SMM_KMS_PM_Master_Document_template.xlsx"
    readiness = root / "working" / "readiness" / "SMM_KMS_PM_Master_Document.xlsx"
    backup = root / "working" / "smm-backups" / "SMM_KMS_PM_Master_Document_2026-07-01_to_2026-07-13.xlsx"
    return ecm, smm, template, readiness, backup


def _hash(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def test_step2_aggregates_flags_hidden_rows_mattex_and_project_alias(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    ecm, _, _, _, _ = _v3_paths(tmp_path, monkeypatch)
    rows = (
        _ecm_row(1007, "Y", "Y", "Y", "Added by Mattex"),
        _ecm_row("1016 WSD", "Y", "Y", "N", "ADD BY CHRIS’S   TEAM"),
        _ecm_row(101, "Y", "N", ""),
        _ecm_row(None, "N", "invalid", "anything", "Added by Mattex"),
    )
    _make_ecm(ecm, rows)
    metrics, displays = aggregate_ecm_record(ecm, ECM_SHEET)
    assert set(metrics) == {101, 1016}
    assert metrics[1016].draft_via_kms == 1
    assert metrics[1016].draft_not_via_kms == 1
    assert metrics[1016].outbound_ecm == 0
    assert metrics[101].not_using_smm == 1
    assert displays == {2: "via KMS", 3: "w/o KMS", 4: "", 5: ""}


@pytest.mark.parametrize(
    ("use_smm", "use_kms", "expected"),
    [("Y", "Y", (1, 0, 1)), ("Y", "N", (0, 1, 1)), ("N", "N", (0, 0, 1)), ("N", None, (0, 0, 1))],
)
def test_step2_validates_smm_kms_combinations(tmp_path: Path, monkeypatch: pytest.MonkeyPatch, use_smm: str, use_kms: str | None, expected: tuple[int, int, int]) -> None:
    ecm, _, _, _, _ = _v3_paths(tmp_path, monkeypatch)
    _make_ecm(ecm, (_ecm_row(101, "Y", use_smm, use_kms),))
    metrics, _ = aggregate_ecm_record(ecm, ECM_SHEET)
    actual = metrics[101]
    assert (actual.draft_via_kms, actual.draft_not_via_kms, actual.raw_outbound_ecm) == expected


@pytest.mark.parametrize(
    ("row", "error"),
    [
        (_ecm_row(101, "Y", None, "Y"), "Use SMM"),
        (_ecm_row(101, "Y", "N", "Y"), "cannot be Y"),
        (_ecm_row(101, "Y", "Y", None), "Use KMS"),
        (_ecm_row(None, "Y", "Y", "Y"), "Project Code is required"),
        (_ecm_row(101, "?", "Y", "Y"), r"Count \(Y/N\)"),
        (_ecm_row(101, "Y", "N", "N", "Added by Mattex"), "Mattex-added Remark"),
    ],
)
def test_step2_fails_closed_on_invalid_rows(tmp_path: Path, monkeypatch: pytest.MonkeyPatch, row: tuple[object, ...], error: str) -> None:
    ecm, _, _, _, _ = _v3_paths(tmp_path, monkeypatch)
    _make_ecm(ecm, (row,))
    with pytest.raises(Step2Error, match=error):
        aggregate_ecm_record(ecm, ECM_SHEET)


def test_step2_rejects_layout_collision_and_preserves_every_source_on_validation_error(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    ecm, smm, template, readiness, backup = _v3_paths(tmp_path, monkeypatch)
    colliding_layout = summary_sheet_name(date(2026, 7, 1), date(2026, 7, 12))
    _make_ecm(ecm, (_ecm_row(101),))
    _make_smm(template, layout_title=colliding_layout)
    _make_smm(smm, layout_title=colliding_layout)
    _make_readiness(readiness)
    smm_before, ecm_before, template_before, readiness_before = _hash(smm), _hash(ecm), _hash(template), _hash(readiness)
    with pytest.raises(Step2Error, match="cannot replace the configured summary layout source"):
        run_step2(ecm, ECM_SHEET, smm, template, readiness, backup, date(2026, 7, 1), date(2026, 7, 12), colliding_layout)
    assert _hash(smm) == smm_before
    assert _hash(ecm) == ecm_before
    assert _hash(template) == template_before
    assert _hash(readiness) == readiness_before
    assert not backup.exists()


def test_step2_transaction_writes_display_summary_postsave_backup_and_reruns(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    ecm, smm, template, readiness, backup = _v3_paths(tmp_path, monkeypatch)
    rows = (
        _ecm_row(101, "Y", "Y", "Y"),
        _ecm_row(101, "Y", "Y", "N"),
        _ecm_row(101, "Y", "N", None),
        _ecm_row(1016, "Y", "Y", "Y", "Added by Mattex"),
        _ecm_row(200, "N", "bad", "bad"),
    )
    _make_ecm(ecm, rows)
    _make_smm(template)
    _make_readiness(readiness)
    _make_smm(smm, readiness=())
    template_hash, snapshot = _hash(template), ecm.parent / "step1b-after.xlsx"
    shutil.copy2(ecm, snapshot)
    snapshot_hash = _hash(snapshot)

    name, count = run_step2(ecm, ECM_SHEET, smm, template, readiness, backup, date(2026, 7, 1), date(2026, 7, 13), LAYOUT_SHEET)
    assert name == SUMMARY_SHEET
    assert count == 2
    assert _hash(template) == template_hash
    assert _hash(snapshot) == snapshot_hash
    assert _hash(backup) == _hash(smm)

    ecm_wb = load_workbook(ecm, data_only=False)
    try:
        sheet = ecm_wb[ECM_SHEET]
        assert [sheet.cell(row, 14).value for row in range(2, 7)] == ["via KMS", "w/o KMS", None, "via KMS", None]
        assert ecm_wb["Unrelated"]["A1"].value == "preserve me"
    finally:
        ecm_wb.close()

    smm_wb = load_workbook(smm, data_only=False)
    try:
        assert smm_wb.active.title == name
        summary = smm_wb[name]
        assert summary["B5"].value == "101 Project 101"
        assert summary["C5"].value == "BU-101"
        assert (summary["E5"].value, summary["F5"].value, summary["G5"].value, summary["H5"].value) == (1, 1, "=E5+F5", 3)
        assert summary["B6"].value == "1016 - 1/WSD/19 (1007)"
        assert summary["H6"].value == 0
        assert summary.column_dimensions["D"].hidden
        assert summary.sheet_view.zoomScale == 85
        assert summary["F5"].font.color.rgb == "FFFF0000"
        assert summary["F6"].font.color.rgb == "FF000000"
        assert smm_wb[LAYOUT_SHEET]["A1"].value == "layout preserved"
        assert smm_wb["Unrelated"]["B2"].value == "=1+1"
    finally:
        smm_wb.close()

    rerun_name, _ = run_step2(ecm, ECM_SHEET, smm, template, readiness, backup, date(2026, 7, 1), date(2026, 7, 13), LAYOUT_SHEET)
    rerun_wb = load_workbook(smm, data_only=False)
    try:
        assert rerun_name == name
        assert rerun_wb.sheetnames.count(name) == 1
        assert rerun_wb.active.title == name
    finally:
        rerun_wb.close()
    assert _hash(backup) == _hash(smm)


def test_step2_initializes_v3_smm_from_template_and_year_names_are_distinct(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    ecm, smm, template, readiness, backup = _v3_paths(tmp_path, monkeypatch)
    _make_ecm(ecm, (_ecm_row(101),), extra_sheet=False)
    _make_smm(template)
    _make_readiness(readiness)
    template_hash = _hash(template)
    run_step2(ecm, ECM_SHEET, smm, template, readiness, backup, date(2026, 7, 1), date(2026, 7, 13), LAYOUT_SHEET)
    assert smm.is_file() and _hash(backup) == _hash(smm)
    assert _hash(template) == template_hash
    assert summary_sheet_name(date(2026, 7, 1), date(2026, 7, 13)) != summary_sheet_name(date(2027, 7, 1), date(2027, 7, 13))


def test_step2_commit_failure_restores_both_workbooks_and_prior_backup(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from control_panel_v3.workers import ecm_step2

    ecm, smm, template, readiness, backup = _v3_paths(tmp_path, monkeypatch)
    _make_ecm(ecm, (_ecm_row(101),), extra_sheet=False)
    _make_smm(template)
    _make_readiness(readiness)
    _make_smm(smm, readiness=())
    backup.parent.mkdir(parents=True, exist_ok=True)
    backup.write_bytes(b"previous-backup-evidence")
    before = {path: _hash(path) for path in (ecm, smm, template, readiness, backup)}
    real_replace = ecm_step2.os.replace
    failure_target = smm

    def fail_smm_publish(source: str | Path, destination: str | Path) -> None:
        if Path(destination) == failure_target and ".v3-step2-smm-" in Path(source).name:
            raise PermissionError("injected SMM publish failure")
        real_replace(source, destination)

    monkeypatch.setattr(ecm_step2.os, "replace", fail_smm_publish)
    with pytest.raises(PermissionError, match="injected SMM publish failure"):
        run_step2(ecm, ECM_SHEET, smm, template, readiness, backup, date(2026, 7, 1), date(2026, 7, 13), LAYOUT_SHEET)
    assert {path: _hash(path) for path in (ecm, smm, template, readiness, backup)} == before


def test_step2_first_run_verification_failure_removes_new_smm_and_restores_ecm(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from control_panel_v3.workers import ecm_step2

    ecm, smm, template, readiness, backup = _v3_paths(tmp_path, monkeypatch)
    _make_ecm(ecm, (_ecm_row(101),), extra_sheet=False)
    _make_smm(template)
    _make_readiness(readiness)
    ecm_before = _hash(ecm)

    def fail_verification(*_args: object) -> None:
        raise Step2Error("injected post-publish verification failure")

    monkeypatch.setattr(ecm_step2, "_verify_active_and_backup", fail_verification)
    with pytest.raises(Step2Error, match="injected post-publish verification failure"):
        run_step2(ecm, ECM_SHEET, smm, template, readiness, backup, date(2026, 7, 1), date(2026, 7, 13), LAYOUT_SHEET)
    assert _hash(ecm) == ecm_before
    assert not smm.exists()
    assert not backup.exists()


def test_step2_rejects_open_workbook_lock_without_mutation(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    ecm, smm, template, readiness, backup = _v3_paths(tmp_path, monkeypatch)
    _make_ecm(ecm, (_ecm_row(101),), extra_sheet=False)
    _make_smm(template)
    _make_readiness(readiness)
    _make_smm(smm)
    before = (_hash(ecm), _hash(smm), _hash(template), _hash(readiness))
    lock = smm.with_name(f"~${smm.name}")
    lock.write_text("locked", encoding="utf-8")
    try:
        with pytest.raises(Step2Error, match="Close the V3 SMM workbook"):
            run_step2(ecm, ECM_SHEET, smm, template, readiness, backup, date(2026, 7, 1), date(2026, 7, 13), LAYOUT_SHEET)
    finally:
        lock.unlink(missing_ok=True)
    assert (_hash(ecm), _hash(smm), _hash(template), _hash(readiness)) == before


def test_step2_builder_is_v3_scoped_and_requires_v3_workbooks(tmp_path: Path) -> None:
    settings = load_settings()
    roots = {name: settings.project_root / "control_panel_v3" / "tests" / f"step2-{name}" for name in (
        "data_dir", "archive_dir", "csv_dir", "ecm_output_dir", "smm_output_dir", "django_output_dir", "edge_profile_dir"
    )}
    readiness = settings.project_root / "control_panel_v3" / "tests" / "step2-readiness" / "readiness.xlsx"
    test_settings = type(settings)(**{**settings.__dict__, **roots, "readiness_workbook_path": readiness, "summary_layout_sheet": LAYOUT_SHEET})
    ecm = test_settings.ecm_output_dir / "ECM_Record.xlsx"
    smm = test_settings.smm_output_dir / "SMM_KMS_PM_Master_Document.xlsx"
    _make_ecm(ecm, (_ecm_row(101),), extra_sheet=False)
    _make_readiness(readiness)
    _make_smm(smm, readiness=())
    prepared = prepare_workflow("ecm-step-2", date(2026, 7, 1), date(2026, 7, 13), 802, test_settings)
    assert "control_panel_v3.workers.ecm_step2" in prepared.command
    assert str(ecm) in prepared.command and str(smm) in prepared.command
    assert "--readiness-workbook" in prepared.command and str(readiness) in prepared.command
    assert "ecm-step-2" not in str(prepared.working_directory)
    assert len(prepared.artifacts) == 3
    assert all(item.path.is_relative_to(CONTROL_PANEL_ROOT) for item in prepared.artifacts)
    assert prepared.artifacts[2].path.is_relative_to(CONTROL_PANEL_ROOT / "working" / "smm-backups")
    shutil.rmtree(test_settings.data_dir, ignore_errors=True)
    shutil.rmtree(test_settings.archive_dir, ignore_errors=True)
    shutil.rmtree(test_settings.csv_dir, ignore_errors=True)
    shutil.rmtree(test_settings.ecm_output_dir, ignore_errors=True)
    shutil.rmtree(test_settings.smm_output_dir, ignore_errors=True)
    shutil.rmtree(test_settings.django_output_dir, ignore_errors=True)
    shutil.rmtree(test_settings.edge_profile_dir, ignore_errors=True)
    shutil.rmtree(readiness.parent, ignore_errors=True)


def test_step2_gate_allows_bypass_only_when_step1b_was_never_run(tmp_path: Path) -> None:
    history = HistoryStore(tmp_path / "history.sqlite3", tmp_path / "runs")
    history.initialize()
    period = history.create_report_period("2026-07-01", "2026-07-13")
    step0 = history.create_run("ecm-step-0", "Step 0", period.start_date, period.end_date, period_id=period.id)
    history.mark_finished(step0.id, state="succeeded", exit_code=0)
    step1a = history.create_run("ecm-step-1a", "Step 1a", period.start_date, period.end_date, period_id=period.id)
    history.mark_finished(step1a.id, state="succeeded", exit_code=0)
    lookup = history.create_run("ecm-first-date-lookup", "Lookup", period.start_date, period.end_date, period_id=period.id)
    history.mark_finished(lookup.id, state="succeeded", exit_code=0)
    summary = history.list_report_period_summaries(REPORT_WORKFLOW_IDS)[0]
    assert ecm_workflow_gate_error(summary, "ecm-step-2", confirm_step_1b_bypass=True) is None
    step1b = history.create_run("ecm-step-1b", "Step 1b", period.start_date, period.end_date, period_id=period.id)
    history.mark_finished(step1b.id, state="failed", exit_code=1, error_summary="Step 1b failed")
    summary = history.list_report_period_summaries(REPORT_WORKFLOW_IDS)[0]
    assert "Step 1b must succeed" in (ecm_workflow_gate_error(summary, "ecm-step-2", confirm_step_1b_bypass=True) or "")
    history.mark_finished(step1b.id, state="succeeded", exit_code=0)
    summary = history.list_report_period_summaries(REPORT_WORKFLOW_IDS)[0]
    assert ecm_workflow_gate_error(summary, "ecm-step-2") is None


def test_step2_bypass_skips_all_ecm_prerequisites_only_when_step1b_was_never_run(tmp_path: Path) -> None:
    history = HistoryStore(tmp_path / "history.sqlite3", tmp_path / "runs")
    history.initialize()
    period = history.create_report_period("2026-07-01", "2026-07-13")
    summary = history.list_report_period_summaries(REPORT_WORKFLOW_IDS)[0]
    assert ecm_workflow_gate_error(summary, "ecm-step-2", confirm_step_1b_bypass=True) is None
    assert "Step 1b must succeed" in (ecm_workflow_gate_error(summary, "ecm-step-2") or "")

    step0 = history.create_run("ecm-step-0", "Step 0", period.start_date, period.end_date, period_id=period.id)
    history.mark_finished(step0.id, state="failed", exit_code=1, error_summary="Step 0 failed")
    summary = history.list_report_period_summaries(REPORT_WORKFLOW_IDS)[0]
    assert ecm_workflow_gate_error(summary, "ecm-step-2", confirm_step_1b_bypass=True) is None