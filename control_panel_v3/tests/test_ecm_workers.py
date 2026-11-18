from __future__ import annotations

import csv
import sys
from datetime import date
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from control_panel_v3.services.config import load_settings
from control_panel_v3.services.workflow_builders import PreflightError, prepare_workflow
from control_panel_v3.workers.ecm_outgoing_reference_lookup import (
    COUNT_COLUMN,
    OUTGOING_REFERENCE_HEADER,
    OUTGOING_REFERENCE_COLUMN,
    REMARK_COLUMN,
    WORKBOOK_HEADERS,
    LookupError,
    build_search_rows,
    read_outgoing_reference_map,
    validate_workbook,
)
from control_panel_v3.workers.properties_csv import CSV_HEADERS, CsvValidationError, validate_properties_csv


def _write_csv(path: Path, *, reference: str = "REF/42", modified_on: str = "") -> None:
    row = {
        "Document title": "Method Statement",
        "Document date": "8/1/2026, 12:00 AM",
        "Name": "MS-42",
        "Document revision": "A",
        "Modified By": "Operator",
        "Major Version": "1",
        "SMM ID": "SMM42",
        "Modified On": modified_on,
        "Outgoing ref. no.": reference,
        "Path": "/folder",
        "Link": "https://ecm.example/item/42",
    }
    with path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=CSV_HEADERS)
        writer.writeheader()
        writer.writerow(row)


def _write_workbook(path: Path, *, count: str = "", revision: str = "A") -> str:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "ECM_1-7 Aug (WIP)"
    for column, value in enumerate(WORKBOOK_HEADERS, start=1):
        sheet.cell(1, column).value = value
    sheet.cell(1, REMARK_COLUMN).value = "Remark"
    sheet.cell(1, OUTGOING_REFERENCE_COLUMN).value = OUTGOING_REFERENCE_HEADER
    values = (1016, "Method Statement", date(2026, 8, 1), "MS-42", revision, "Operator", 1, "SMM42", count, None, None)
    for column, value in enumerate(values, start=1):
        sheet.cell(2, column).value = value
    workbook.save(path)
    workbook.close()
    return sheet.title


def test_step0_csv_validator_accepts_contract_and_optional_modified_on(tmp_path: Path) -> None:
    path = tmp_path / "properties.csv"
    _write_csv(path)
    assert validate_properties_csv(path) == 1


def test_step0_csv_validator_rejects_wrong_header_contract(tmp_path: Path) -> None:
    path = tmp_path / "bad.csv"
    path.write_text("Document title,Wrong\n", encoding="utf-8")
    with pytest.raises(CsvValidationError, match="headers"):
        validate_properties_csv(path)


def test_step0_publication_failure_keeps_existing_period_handoff(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    import control_panel_v3.workers.ecm_step0_export as step0

    temporary = tmp_path / "download.csv"
    output = tmp_path / "run.csv"
    handoff = tmp_path / "period.csv"
    _write_csv(temporary)
    handoff.write_text("previous handoff", encoding="utf-8")
    original_replace = step0.os.replace

    def fail_handoff(source, destination):
        if Path(destination) == handoff:
            raise PermissionError("injected handoff publish failure")
        original_replace(source, destination)

    monkeypatch.setattr(step0.os, "replace", fail_handoff)
    with pytest.raises(PermissionError, match="injected handoff publish failure"):
        step0._publish_validated_exports(temporary, output, handoff)
    assert not output.exists()
    assert handoff.read_text(encoding="utf-8") == "previous handoff"
    assert not output.with_name(f".{output.name}.publish").exists()
    assert not handoff.with_name(f".{handoff.name}.publish").exists()


def test_lookup_uses_exact_csv_reference_and_skips_count_n_or_blank_revision(tmp_path: Path) -> None:
    csv_path = tmp_path / "properties.csv"
    workbook_path = tmp_path / "ECM_Record.xlsx"
    _write_csv(csv_path)
    sheet_name = _write_workbook(workbook_path)
    workbook = load_workbook(workbook_path)
    try:
        sheet = workbook[sheet_name]
        ref_map = read_outgoing_reference_map(csv_path)
        tasks = build_search_rows(sheet, ref_map)
        assert tasks == [(2, "REF/42")]

        sheet.cell(2, COUNT_COLUMN).value = "N"
        assert build_search_rows(sheet, ref_map) == []
        sheet.cell(2, COUNT_COLUMN).value = None
        sheet.cell(2, 5).value = None
        assert build_search_rows(sheet, ref_map) == []
    finally:
        workbook.close()


def test_lookup_rerun_scopes_preserve_populated_remarks(tmp_path: Path) -> None:
    csv_path = tmp_path / "properties.csv"
    workbook_path = tmp_path / "ECM_Record.xlsx"
    _write_csv(csv_path)
    sheet_name = _write_workbook(workbook_path)
    workbook = load_workbook(workbook_path)
    try:
        sheet = workbook[sheet_name]
        sheet.cell(2, REMARK_COLUMN).value = "existing hit"
        references = read_outgoing_reference_map(csv_path)
        assert build_search_rows(sheet, references, scope="all") == [(2, "REF/42")]
        assert build_search_rows(sheet, references, scope="blank") == []
        assert build_search_rows(sheet, references, scope="previously-unmatched", previously_unmatched={2}) == []
        sheet.cell(2, REMARK_COLUMN).value = None
        assert build_search_rows(sheet, references, scope="previously-unmatched", previously_unmatched={2}) == [(2, "REF/42")]
        with pytest.raises(LookupError, match="prior lookup summary"):
            build_search_rows(sheet, references, scope="previously-unmatched")
    finally:
        workbook.close()


def test_lookup_refuses_ambiguous_csv_identity(tmp_path: Path) -> None:
    path = tmp_path / "properties.csv"
    _write_csv(path, reference="REF/42")
    with path.open("a", encoding="utf-8", newline="") as stream:
        row = ["Method Statement", "8/1/2026, 12:00 AM", "MS-42", "A", "Operator", "1", "SMM42", "", "OTHER", "/folder", "https://ecm.example/item/42"]
        csv.writer(stream).writerow(row)
    with pytest.raises(LookupError, match="ambiguous"):
        read_outgoing_reference_map(path)


def test_lookup_does_not_normalize_document_title_as_project_code() -> None:
    from control_panel_v3.workers.ecm_outgoing_reference_lookup import _identity

    assert _identity(("1007", "2026-08-01", "MS-42", "A", "Operator", "1", "SMM42"))[0] == "1007"


def test_lookup_workbook_requires_v3_contract(tmp_path: Path) -> None:
    path = tmp_path / "ECM_Record.xlsx"
    sheet_name = _write_workbook(path)
    workbook = load_workbook(path)
    try:
        sheet = workbook[sheet_name]
        sheet.cell(1, OUTGOING_REFERENCE_COLUMN).value = "Wrong"
        with pytest.raises(LookupError, match="column P"):
            validate_workbook(sheet)
    finally:
        workbook.close()


def test_lookup_enforces_out_and_starts_with_criteria(monkeypatch: pytest.MonkeyPatch) -> None:
    import asyncio

    import control_panel_v3.workers.ecm_outgoing_reference_lookup as lookup

    class Control:
        def __init__(self, value: str = "", *, control_id: str = "", class_name: str = "") -> None:
            self.value = value
            self.control_id = control_id
            self.class_name = class_name
            self.filled = None

        async def input_value(self):
            return self.value

        async def fill(self, value):
            self.filled = value

        async def dispatch_event(self, *_args):
            return None

        async def press(self, *_args):
            return None

    direction = Control("OUT")
    operator = Control("Starts With", control_id="OutgoingRefOperator")
    reference = Control()

    async def labeled(_form, label, *, include_operator=False, **_kwargs):
        if label == "In/Out":
            return [direction]
        return [operator, reference] if include_operator else [reference]

    async def metadata(control):
        return {"id": control.control_id, "class_name": control.class_name}

    monkeypatch.setattr(lookup, "_labeled_controls", labeled)
    monkeypatch.setattr(lookup, "_control_metadata", metadata)
    asyncio.run(lookup._set_search_criteria(object(), "EXACT/REF"))
    assert reference.filled == "EXACT/REF"


def test_lookup_reads_oldest_date_across_result_pages(monkeypatch: pytest.MonkeyPatch) -> None:
    import asyncio

    import control_panel_v3.workers.ecm_outgoing_reference_lookup as lookup

    class Text:
        def __init__(self, text: str) -> None:
            self.text = text

        async def inner_text(self):
            return self.text

    class Cell(Text):
        pass

    class Row(Text):
        def __init__(self, text: str, cells: list[Cell]) -> None:
            super().__init__(text)
            self.cells = cells

        def get_by_role(self, role, **_kwargs):
            assert role == "gridcell"
            return Collection(self.cells)

    class Collection:
        def __init__(self, items):
            self.items = items
            self.first = items[0] if items else None

        async def all(self):
            return self.items

        async def count(self):
            return len(self.items)

        def nth(self, index):
            return self.items[index]

    class NextButton:
        def __init__(self, grid):
            self.grid = grid

        async def is_enabled(self):
            return self.grid.page_number == 0

        async def get_attribute(self, _name):
            return None

        async def click(self, **_kwargs):
            self.grid.page_number += 1

    class Toolbar:
        def __init__(self, grid):
            self.grid = grid
            self.first = self

        async def wait_for(self, **_kwargs):
            return None

        def get_by_role(self, role, **_kwargs):
            assert role == "button"
            return Collection([NextButton(self.grid)])

    class Grid:
        def __init__(self):
            self.page_number = 0

        def get_by_role(self, role, **_kwargs):
            if role == "columnheader":
                return Collection([Text("Document date")])
            if role == "row":
                date_value = "8/3/2026, 12:00 AM" if self.page_number == 0 else "8/1/2026, 12:00 AM"
                return Collection([Row("header", []), Row(date_value, [Cell(date_value)])])
            raise AssertionError(role)

    class Panel:
        def __init__(self):
            self.grid = Grid()

        def get_by_role(self, role, **_kwargs):
            if role == "toolbar":
                return Toolbar(self.grid)
            if role == "grid":
                return self.grid
            raise AssertionError(role)

    class Page:
        def __init__(self):
            self.panel = None

        def locator(self, _selector):
            return self

        async def inner_text(self):
            return "The search returned 2 items"

    monkeypatch.setattr(lookup, "_visible_locators", lambda locator: asyncio.sleep(0, result=locator.items))
    panel, page = Panel(), Page()
    result = asyncio.run(lookup._read_oldest_result_date(panel, page, 1000))
    assert result == date(2026, 8, 1)


def test_lookup_processes_rows_independently_and_leaves_misses_unchanged(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    import asyncio
    import json
    import types

    import control_panel_v3.workers.ecm_outgoing_reference_lookup as lookup

    root = Path(lookup.__file__).resolve().parents[1]
    csv_path = root / "tests" / "lookup-fixture.csv"
    workbook_path = root / "tests" / "lookup-fixture.xlsx"
    profile_path = root / "tests" / "lookup-profile"
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    rows = [
        ["Method Statement", "8/1/2026, 12:00 AM", "MS-42", "A", "Operator", "1", "SMM42", "", "REF/HIT", "/folder", "https://ecm.example/item/42"],
        ["Method Statement", "8/2/2026, 12:00 AM", "MS-43", "A", "Operator", "1", "SMM43", "", "REF/MISS", "/folder", "https://ecm.example/item/43"],
    ]
    with csv_path.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(CSV_HEADERS)
        writer.writerows(rows)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Lookup"
    for column, value in enumerate(WORKBOOK_HEADERS, start=1):
        sheet.cell(1, column).value = value
    sheet.cell(1, REMARK_COLUMN).value = "Remark"
    sheet.cell(1, OUTGOING_REFERENCE_COLUMN).value = OUTGOING_REFERENCE_HEADER
    for row_number, values in enumerate(((1016, "Method Statement", date(2026, 8, 1), "MS-42", "A", "Operator", 1, "SMM42", "", None, None), (1016, "Method Statement", date(2026, 8, 2), "MS-43", "A", "Operator", 1, "SMM43", "", None, None)), start=2):
        for column, value in enumerate(values, start=1):
            sheet.cell(row_number, column).value = value
    sheet.cell(3, REMARK_COLUMN).value = "reviewed manually"
    workbook.save(workbook_path)
    workbook.close()

    class FakeGrid:
        def __init__(self, current_page: int = 0):
            self.current_page = current_page

        async def get_by_role(self, role, name=None):
            if role == "columnheader":
                return FakeCollection([FakeText("Document date")])
            if role == "row":
                if current_reference[0] == "REF/MISS":
                    return FakeCollection([FakeRow("header", ()), FakeRow("result", (FakeText(""),))])
                dates = (FakeText("8/4/2026, 12:00 AM"), FakeText("8/1/2026, 12:00 AM"))
                return FakeCollection([FakeRow("header", ()), FakeRow("result", (dates[0 if self.current_page == 0 else 1],))])
            raise AssertionError(role)

        async def locator(self, selector):
            return self

    class FakeText:
        def __init__(self, text):
            self.text = text

        async def inner_text(self):
            return self.text

    class FakeCollection:
        def __init__(self, values):
            self.values = values

        async def all(self):
            return self.values

        async def count(self):
            return len(self.values)

        def nth(self, index):
            return self.values[index]

        async def is_visible(self):
            return True

    class FakeRow(FakeText):
        def __init__(self, text, cells):
            super().__init__(text)
            self.cells = cells

        def get_by_role(self, role, name=None, **kwargs):
            return FakeCollection(self.cells)

    class FakeToolbar:
        def get_by_role(self, role, name=None, **kwargs):
            return FakeCollection([])

        async def wait_for(self, **kwargs):
            return None

    current_reference = [""]
    fail_result_read = [False]

    class FakePanel:
        def __init__(self):
            self.page_number = 0

        def get_by_role(self, role, name=None, **kwargs):
            return FakeGrid(self.page_number) if role == "grid" else FakeToolbar()

    class FakePage:
        def __init__(self):
            self.panel = FakePanel()

        def locator(self, selector):
            return self

        async def goto(self, *args, **kwargs):
            return None

        def get_by_role(self, role, name=None):
            return FakeButtonCollection()

        async def inner_text(self):
            return "The search returned 0 items" if current_reference[0] == "REF/MISS" else "The search returned 2 items"

    class FakeButtonCollection(FakeCollection):
        async def click(self, **kwargs):
            return None

        async def is_enabled(self):
            return True

        async def get_attribute(self, name):
            return None

    class FakeContext:
        def __init__(self):
            self.pages = [FakePage()]

        async def close(self):
            return None

    class FakeBrowser:
        async def launch_persistent_context(self, **kwargs):
            return FakeContext()

    class FakePlaywright:
        chromium = FakeBrowser()

        async def __aenter__(self):
            return self

        async def __aexit__(self, *args):
            return None

    class FakePlaywrightFactory:
        async def __aenter__(self):
            return FakePlaywright()

        async def __aexit__(self, *args):
            return None

    monkeypatch.setitem(sys.modules, "playwright.async_api", types.SimpleNamespace(async_playwright=lambda: FakePlaywrightFactory()))

    async def fake_open_saved_search(page, timeout, **kwargs):
        return page.panel

    async def fake_active_form(panel):
        return panel

    async def fake_set_criteria(form, reference):
        current_reference[0] = reference

    async def fake_visible(locator):
        return [FakeButton()]

    class FakeButton:
        async def click(self, **kwargs):
            return None

    monkeypatch.setattr(lookup, "_open_saved_search", fake_open_saved_search)
    monkeypatch.setattr(lookup, "_active_form", fake_active_form)
    monkeypatch.setattr(lookup, "_set_search_criteria", fake_set_criteria)
    monkeypatch.setattr(lookup, "_visible_locators", fake_visible)
    async def fake_read(panel, page, timeout):
        if fail_result_read[0]:
            raise TimeoutError("search results did not load")
        if current_reference[0] == "REF/MISS":
            return None
        return date(2026, 8, 1)

    monkeypatch.setattr(lookup, "_read_oldest_result_date", fake_read)
    monkeypatch.setattr(lookup, "asyncio", asyncio)
    summary_path = root / "tests" / "lookup-summary.json"
    try:
        fail_result_read[0] = True
        with pytest.raises(LookupError, match="search results did not load"):
            asyncio.run(lookup.run_lookup(workbook_path=workbook_path, worksheet_name="Lookup", csv_path=csv_path, profile_dir=profile_path, ecm_url="https://example.test/", timeout_seconds=2, summary_path=summary_path))
        assert not summary_path.exists()
        fail_result_read[0] = False
        result = asyncio.run(lookup.run_lookup(workbook_path=workbook_path, worksheet_name="Lookup", csv_path=csv_path, profile_dir=profile_path, ecm_url="https://example.test/", timeout_seconds=2, summary_path=summary_path))
        assert result == (1, 1, 2)
        summary = json.loads(summary_path.read_text(encoding="utf-8"))
        assert summary["partial_result"] is True
        assert summary["miss_count"] == 1
        assert summary["misses"] == [{"row": 3, "outgoing_reference": "REF/MISS"}]
        updated = load_workbook(workbook_path)
        try:
            assert updated["Lookup"].cell(2, REMARK_COLUMN).value == "2026-08-01"
            assert updated["Lookup"].cell(3, REMARK_COLUMN).value == "reviewed manually"
        finally:
            updated.close()
    finally:
        for path in (csv_path, workbook_path, summary_path):
            path.unlink(missing_ok=True)
        profile_path.rmdir() if profile_path.exists() and not any(profile_path.iterdir()) else None


def test_worker_builder_creates_only_v3_step0_command(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    settings = load_settings()
    settings.csv_dir.mkdir(parents=True, exist_ok=True)
    prepared = prepare_workflow("ecm-step-0", date(2026, 8, 1), date(2026, 8, 7), 42, settings)
    assert "control_panel_v3.workers.ecm_step0_export" in prepared.command
    assert all("control_panel_v2" not in argument and "ecm\\" not in argument.casefold() for argument in prepared.command)
    assert prepared.working_directory == settings.project_root
    assert all(item.path.is_relative_to(settings.csv_dir) for item in prepared.artifacts[:2])
    assert all(item.path.is_relative_to(settings.runs_dir) for item in prepared.artifacts[2:])
    assert settings.edge_profile_dir.is_relative_to(settings.project_root / "control_panel_v3")


def test_worker_builder_uses_unique_timestamped_output_and_stable_period_handoff() -> None:
    settings = load_settings()
    first = prepare_workflow("ecm-step-0", date(2026, 8, 1), date(2026, 8, 7), 43, settings)
    second = prepare_workflow("ecm-step-0", date(2026, 8, 1), date(2026, 8, 7), 44, settings)
    assert first.artifacts[0].path != second.artifacts[0].path
    assert first.artifacts[1].path == second.artifacts[1].path


def test_step0_worker_rejects_caller_selected_non_v3_root(tmp_path: Path) -> None:
    import asyncio

    from control_panel_v3.workers.ecm_step0_export import Step0Error, run_export

    with pytest.raises(Step0Error, match="root is fixed"):
        asyncio.run(run_export(
            start_date=date(2026, 8, 1), end_date=date(2026, 8, 7),
            output_path=tmp_path / "outside.csv", period_input_path=None,
            ecm_url="https://example.test/", profile_dir=tmp_path / "profile",
            timeout_seconds=1, root_dir=tmp_path,
        ))


def test_browser_support_resets_profile_to_timestamped_backup(tmp_path: Path) -> None:
    from control_panel_v3.workers.browser_support import reset_profile

    profile = tmp_path / "edge-profile"
    profile.mkdir()
    (profile / "session.txt").write_text("local session", encoding="utf-8")
    backup = reset_profile(profile)
    assert backup is not None and backup.name.startswith("edge-profile.backup-")
    assert not profile.exists()
    assert (backup / "session.txt").read_text(encoding="utf-8") == "local session"


def test_browser_support_diagnostics_strip_url_query_and_fragment() -> None:
    from control_panel_v3.workers.browser_support import BrowserDiagnostics

    class Page:
        def __init__(self):
            self.handlers = {}

        def on(self, event, callback):
            self.handlers[event] = callback

    class Request:
        method = "GET"
        resource_type = "document"
        url = "https://example.test/path?token=secret#fragment"

    class Response:
        status = 500
        url = "https://example.test/path?password=secret#fragment"
        request = Request()

    page, diagnostics = Page(), BrowserDiagnostics()
    diagnostics.attach(page)
    page.handlers["requestfailed"](Request())
    page.handlers["response"](Response())
    payload = str(diagnostics.events)
    assert "token=secret" not in payload and "password=secret" not in payload
    assert "fragment" not in payload


def test_browser_support_cursor_is_page_scoped_and_pauses_are_bounded(monkeypatch: pytest.MonkeyPatch) -> None:
    import asyncio
    import random

    from control_panel_v3.workers import browser_support

    pauses = []

    async def fake_sleep(value):
        pauses.append(value)

    class Mouse:
        def __init__(self):
            self.points = []

        async def move(self, x, y):
            self.points.append((x, y))

    class Page:
        def __init__(self):
            self.mouse = Mouse()

    class Locator:
        async def bounding_box(self):
            return {"x": 100, "y": 100, "width": 40, "height": 20}

    monkeypatch.setattr(browser_support.asyncio, "sleep", fake_sleep)
    page = Page()
    moved = asyncio.run(browser_support.PageCursor(page, rng=random.Random(7)).move_to(Locator()))
    assert moved is True
    assert len(page.mouse.points) >= 8
    assert pauses and max(pauses) <= 1.5


def test_browser_support_waits_for_manual_authentication(monkeypatch: pytest.MonkeyPatch) -> None:
    import asyncio

    from control_panel_v3.workers.browser_support import wait_for_manual_auth

    authenticated = [False]
    ready_calls = []

    async def fake_sleep(_value):
        authenticated[0] = True

    async def is_authenticated(_page):
        return authenticated[0]

    async def ready(_page, timeout):
        ready_calls.append(timeout)

    monkeypatch.setattr("control_panel_v3.workers.browser_support.asyncio.sleep", fake_sleep)
    asyncio.run(wait_for_manual_auth(object(), is_authenticated, ready, 1))
    assert ready_calls == [2_000]


def test_app_routes_show_only_supported_live_worker_controls(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from fastapi.testclient import TestClient

    from control_panel_v3 import main as app_module

    settings = load_settings()
    route_root = settings.project_root / "control_panel_v3" / "tests" / f"route-{tmp_path.name}"
    settings = type(settings)(**{**settings.__dict__, "data_dir": route_root / "data", "archive_dir": route_root / "archive", "csv_dir": route_root / "csv", "ecm_output_dir": route_root / "ecm", "smm_output_dir": route_root / "smm", "django_output_dir": route_root / "django", "edge_profile_dir": route_root / "profile", "weekly_template_path": route_root / "templates" / "weekly.xlsx"})
    settings.ensure_application_directories()
    history = app_module.HistoryStore(settings.database_path, settings.runs_dir)
    history.initialize()
    history.create_report_period("2026-08-01", "2026-08-07")
    monkeypatch.setattr(app_module, "load_settings", lambda: settings)
    app = app_module.create_app()
    app.state.preview_only = False
    client = TestClient(app)
    response = client.get("/report-history")
    assert response.status_code == 200
    assert 'name="workflow_id"' not in response.text
    assert "Step 2" in response.text
    assert "manual Use SMM / Use KMS review" in response.text
    assert "Weekly reporting is independent of ECM gates" in response.text
    assert "Weekly Method Statement Report" in response.text
    client.close()
    directories = (route_root,)
    import shutil

    for directory in directories:
        if directory.exists():
            shutil.rmtree(directory, ignore_errors=True)


def test_report_workflow_routes_enforce_preview_handoff_and_audit_metadata(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from fastapi.testclient import TestClient

    from control_panel_v3 import main as app_module

    settings = load_settings()
    test_root = settings.project_root / "control_panel_v3" / "tests" / f"routes-{tmp_path.name}"
    roots = {name: test_root / name for name in (
        "data_dir", "archive_dir", "csv_dir", "ecm_output_dir", "smm_output_dir", "django_output_dir", "edge_profile_dir"
    )}
    test_settings = type(settings)(**{**settings.__dict__, **roots})
    monkeypatch.setattr(app_module, "load_settings", lambda: test_settings)
    app = app_module.create_app()
    history = app.state.history
    client = TestClient(app)
    try:
        period = history.create_report_period("2026-08-01", "2026-08-07")
        preview = client.post(f"/report-periods/{period.id}/runs/ecm-step-2", data={"confirm_step_1b_bypass": "true"}, follow_redirects=False)
        assert preview.status_code == 303
        assert "preview-only" in preview.headers["location"]
        assert history.list_runs(period_id=period.id) == []

        app.state.preview_only = False
        submitted: list[dict[str, object]] = []

        def record_submit(workflow_id, label, start_date, end_date, **kwargs):
            run = history.create_run(
                workflow_id, label, start_date.isoformat(), end_date.isoformat(),
                period_id=kwargs.get("period_id"),
                manual_step_1b_bypass=kwargs.get("manual_step_1b_bypass", False),
                lookup_scope=kwargs.get("lookup_scope"),
            )
            history.mark_finished(run.id, state="succeeded", exit_code=0)
            submitted.append({"workflow_id": workflow_id, **kwargs})
            return history.get_run(run.id)

        monkeypatch.setattr(app.state.runner, "submit", record_submit)
        step0 = history.create_run("ecm-step-0", "ECM Step 0", period.start_date, period.end_date, period_id=period.id)
        history.mark_finished(step0.id, state="succeeded", exit_code=0)
        accepted = client.post(f"/report-periods/{period.id}/runs/next", data={"source_run_id": step0.id}, follow_redirects=False)
        assert accepted.status_code == 303
        assert submitted[-1]["workflow_id"] == "ecm-step-1a"

        stale_period = history.create_report_period("2026-07-25", "2026-07-31")
        old_step0 = history.create_run("ecm-step-0", "ECM Step 0", stale_period.start_date, stale_period.end_date, period_id=stale_period.id)
        history.mark_finished(old_step0.id, state="succeeded", exit_code=0)
        latest_step0 = history.create_run("ecm-step-0", "ECM Step 0", stale_period.start_date, stale_period.end_date, period_id=stale_period.id)
        history.mark_finished(latest_step0.id, state="succeeded", exit_code=0)
        stale_source = client.post(f"/report-periods/{stale_period.id}/runs/next", data={"source_run_id": old_step0.id}, follow_redirects=False)
        assert "no+longer+the+latest" in stale_source.headers["location"]

        lookup_period = history.create_report_period("2026-08-08", "2026-08-14")
        for workflow_id, label in (("ecm-step-0", "Step 0"), ("ecm-step-1a", "Step 1a")):
            prerequisite = history.create_run(workflow_id, label, lookup_period.start_date, lookup_period.end_date, period_id=lookup_period.id)
            history.mark_finished(prerequisite.id, state="succeeded", exit_code=0)
        lookup_response = client.post(
            f"/report-periods/{lookup_period.id}/runs/ecm-first-date-lookup",
            data={"lookup_scope": "blank"}, follow_redirects=False,
        )
        assert lookup_response.status_code == 303
        lookup_run = history.list_runs(workflow_id="ecm-first-date-lookup", period_id=lookup_period.id)[0]
        assert lookup_run.lookup_scope == "blank"

        bypass_period = history.create_report_period("2026-08-15", "2026-08-21")
        bypass_response = client.post(
            f"/report-periods/{bypass_period.id}/runs/ecm-step-2",
            data={"confirm_step_1b_bypass": "true"}, follow_redirects=False,
        )
        assert bypass_response.status_code == 303
        bypass_run = history.list_runs(workflow_id="ecm-step-2", period_id=bypass_period.id)[0]
        assert bypass_run.manual_step_1b_bypass is True
        assert "Manual Step 1b bypass confirmed" in client.get("/history").text
    finally:
        client.close()
        import shutil
        shutil.rmtree(test_root, ignore_errors=True)
