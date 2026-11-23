from __future__ import annotations

import csv
import shutil
from datetime import date
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from control_panel_v3.services.config import CONTROL_PANEL_ROOT, load_settings
from control_panel_v3.services.workflow_builders import PreflightError, prepare_workflow
from control_panel_v3.workers.ecm_step1a import (
    OUTGOING_REFERENCE_HEADER,
    READINESS_SHEET_NAME,
    REVIEW_HEADERS,
    Step1aError,
    apply_preliminary_rules,
    run_step1a,
)
from control_panel_v3.workers.properties_csv import CSV_HEADERS


def _csv(path: Path, rows: list[list[str]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(CSV_HEADERS)
        writer.writerows(rows)


def _row(name: str, major: str = "1", *, modified: str = "8/1/2026, 12:00 AM", document_date: str = "8/1/2026, 12:00 AM", project: str = "CWCMGCJV/1016/CSF") -> list[str]:
    return ["Method Statement", document_date, f"{project}/{name}.pdf", "A", "Operator", major, "SMM-1", modified, f"{project}/{name}.pdf/A", "/folder", "https://ecm.example/item"]


def _template(path: Path, *, seed_rows: int = 1) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "ECM_template"
    for index, header in enumerate(REVIEW_HEADERS, 1):
        sheet.cell(1, index).value = header
    sheet.cell(1, 12).value = "Work Flow"
    sheet.cell(1, 13).value = "Remark"
    sheet.cell(1, 14).value = "Display"
    for col in range(1, 17):
        sheet.cell(2, col).value = "template value"
    for index in range(seed_rows):
        other = workbook.create_sheet(f"Keep{index}")
        other["A1"] = f"preserve-{index}"
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    workbook.close()


def _readiness(path: Path, codes: list[tuple[int, str]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = READINESS_SHEET_NAME
    sheet.cell(1, 3).value = "Project"
    sheet.cell(1, 5).value = "Ready"
    for row, (code, ready) in enumerate(codes, 2):
        sheet.cell(row, 3).value = f"{code} Project"
        sheet.cell(row, 5).value = ready
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    workbook.close()


def test_preliminary_rules_preserve_legacy_duplicates_and_project_alias() -> None:
    from datetime import datetime
    from control_panel_v3.workers.ecm_step1a import EcmRecord

    records = [
        EcmRecord(2, "Doc", datetime(2026, 8, 1), "CWCMGCJV/1007/CSF/Doc.pdf", "A", "User", 1, "", None, "ref", 1007),
        EcmRecord(3, "Doc", datetime(2026, 8, 1), "CWCMGCJV/1007/CSF/Doc.pdf", "A", "User", 3, "", None, "ref", 1007),
        EcmRecord(4, "Doc", datetime(2026, 8, 1), "Unknown.pdf", "A", "User", 1, "", None, "ref"),
    ]
    apply_preliminary_rules(records, {1016: True})
    assert [record.count for record in records] == ["N", "", "N"]
    assert "superseded_major_version" in records[0].reasons
    assert "unresolved_project_code" in records[2].reasons


def test_step1a_consumes_exact_csv_backs_up_and_preserves_unrelated_sheets(tmp_path: Path) -> None:
    csv_path, readiness, workbook_path, template, backup = (tmp_path / name for name in (
        "period.csv", "readiness.xlsx", "ECM_Record.xlsx", "template.xlsx", "backup.xlsx"
    ))
    _csv(csv_path, [_row("Doc", "1"), _row("Doc", "3"), _row("Outside", document_date="8/2/2026, 12:00 AM")])
    _readiness(readiness, [(1016, "Yes")])
    _template(template)
    _template(workbook_path)
    workbook = load_workbook(workbook_path)
    workbook["Keep0"]["A1"] = "period original"
    workbook.save(workbook_path)
    workbook.close()

    result = run_step1a(csv_path, readiness, workbook_path, template, backup, date(2026, 8, 1), date(2026, 8, 1))
    assert result == {"rows_written": 2, "pending_first_dates": 1, "count_n": 1}
    assert backup.is_file()
    sheet_name = "ECM_1-1 Aug (WIP)"
    workbook = load_workbook(workbook_path, data_only=False)
    try:
        assert "Keep0" in workbook.sheetnames
        assert workbook["Keep0"]["A1"].value == "period original"
        sheet = workbook[sheet_name]
        assert sheet.cell(1, 16).value == OUTGOING_REFERENCE_HEADER
        assert sheet.cell(2, 9).value == "N"
        assert sheet.cell(3, 9).value is None
        assert sheet.cell(2, 13).value == "Duplicate with its version 3"
        assert sheet.cell(2, 14).value is None
        assert sheet.cell(2, 16).value == "CWCMGCJV/1016/CSF/Doc.pdf"
        assert sheet.auto_filter.ref == "A1:P3"
    finally:
        workbook.close()


def test_step1a_removes_matching_parenthesized_revision_from_column_p(tmp_path: Path) -> None:
    csv_path, readiness, workbook_path, template, backup = (tmp_path / name for name in (
        "period.csv", "readiness.xlsx", "ECM_Record.xlsx", "template.xlsx", "backup.xlsx"
    ))
    project = "CWCRCC(ND202404)JV/1131/CSF"
    row = _row("MS/0283-2025(F", project=project)
    row[3] = "F"
    row[8] = "CWCRCC(ND202404)JV/1131/CSF/MS/0283-2025(F"
    _csv(csv_path, [row])
    _readiness(readiness, [(1131, "Yes")])
    _template(template)
    _template(workbook_path)

    run_step1a(csv_path, readiness, workbook_path, template, backup, date(2026, 8, 1), date(2026, 8, 1))

    workbook = load_workbook(workbook_path)
    try:
        assert workbook.active.cell(2, 16).value == "CWCRCC(ND202404)JV/1131/CSF/MS/0283-2025"
    finally:
        workbook.close()


def test_step1a_initializes_from_template_and_rejects_bad_input_without_mutation(tmp_path: Path) -> None:
    csv_path, readiness, workbook_path, template, backup = (tmp_path / name for name in (
        "period.csv", "readiness.xlsx", "ECM_Record.xlsx", "template.xlsx", "backup.xlsx"
    ))
    _csv(csv_path, [_row("Doc")])
    _readiness(readiness, [(1016, "Yes")])
    _template(template)
    run_step1a(csv_path, readiness, workbook_path, template, backup, date(2026, 8, 1), date(2026, 8, 1))
    assert workbook_path.is_file() and backup.is_file()

    before = workbook_path.read_bytes()
    bad_template = tmp_path / "bad.xlsx"
    _template(bad_template)
    workbook = load_workbook(bad_template)
    workbook["ECM_template"].cell(1, 1).value = "Wrong"
    workbook.save(bad_template)
    workbook.close()
    second_backup = tmp_path / "second-backup.xlsx"
    with pytest.raises(Step1aError, match="headers"):
        run_step1a(csv_path, readiness, tmp_path / "second-workbook.xlsx", bad_template, second_backup, date(2026, 8, 1), date(2026, 8, 1))
    assert workbook_path.read_bytes() == before


def test_step1a_preflight_uses_exact_v3_handoff_and_v3_owned_paths(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    settings = load_settings()
    roots = {name: settings.project_root / "control_panel_v3" / "tests" / f"step1a-{name}" for name in (
        "data_dir", "archive_dir", "csv_dir", "ecm_output_dir", "smm_output_dir", "django_output_dir", "edge_profile_dir"
    )}
    template_dir = settings.project_root / "control_panel_v3" / "working" / "step1a-test-templates"
    template = template_dir / "ECM_Record_template.xlsx"
    readiness = settings.project_root / "control_panel_v3" / "tests" / "step1a-readiness" / "readiness.xlsx"
    for path in (*roots.values(), template_dir, readiness.parent):
        if path.is_dir():
            shutil.rmtree(path, ignore_errors=True)
        else:
            path.unlink(missing_ok=True)
    test_settings = type(settings)(**{**settings.__dict__, **roots, "ecm_template_path": template, "readiness_workbook_path": readiness})
    test_settings.ensure_application_directories()
    _template(template)
    _readiness(readiness, [(1016, "Yes")])
    period_csv = test_settings.csv_dir / "properties_2026-08-01_to_2026-08-07.csv"
    _csv(period_csv, [_row("Doc")])
    prepared = prepare_workflow("ecm-step-1a", date(2026, 8, 1), date(2026, 8, 7), 501, test_settings)
    assert "control_panel_v3.workers.ecm_step1a" in prepared.command
    assert str(period_csv) in prepared.command
    assert "newest" not in " ".join(prepared.command).casefold()
    assert all(item.path.is_relative_to(CONTROL_PANEL_ROOT) for item in prepared.artifacts)
    assert "run501" in str(prepared.artifacts[1].path)
    assert not any("control_panel_v2" in arg or "ecm\\" in arg.casefold() for arg in prepared.command)
    missing_handoff = test_settings.csv_dir / "properties_2026-08-08_to_2026-08-14.csv"
    with pytest.raises(PreflightError, match="exact-period CSV"):
        prepare_workflow("ecm-step-1a", date(2026, 8, 8), date(2026, 8, 14), 502, test_settings)
    for path in (*roots.values(), template_dir, readiness.parent):
        if path.is_dir():
            shutil.rmtree(path, ignore_errors=True)
        else:
            path.unlink(missing_ok=True)