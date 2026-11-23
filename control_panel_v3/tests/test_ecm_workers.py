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
from control_panel_v3.workers.reference_normalization import outgoing_reference_without_revision


def _write_csv(path: Path, *, reference: str = "REF/42", modified_on: str = "", revision: str = "A") -> None:
    row = {
        "Document title": "Method Statement",
        "Document date": "8/1/2026, 12:00 AM",
        "Name": "MS-42",
        "Document revision": revision,
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
    sheet.cell(2, OUTGOING_REFERENCE_COLUMN).value = "REF/42"
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


def test_lookup_uses_excel_reference_and_skips_count_n_or_blank_revision(tmp_path: Path) -> None:
    csv_path = tmp_path / "properties.csv"
    workbook_path = tmp_path / "ECM_Record.xlsx"
    outgoing_reference = "CWCRCC(ND202404)JV/1131/CSF/MS/0283-2025(F"
    normalized_reference = "CWCRCC(ND202404)JV/1131/CSF/MS/0283-2025"
    _write_csv(csv_path, reference=outgoing_reference, revision="F")
    sheet_name = _write_workbook(workbook_path, revision="F")
    workbook = load_workbook(workbook_path)
    try:
        sheet = workbook[sheet_name]
        ref_map = read_outgoing_reference_map(csv_path)
        assert list(ref_map.values()) == [outgoing_reference]
        tasks = build_search_rows(sheet, ref_map)
        assert tasks == [(2, "REF/42")]
        sheet.cell(2, OUTGOING_REFERENCE_COLUMN).value = outgoing_reference
        assert build_search_rows(sheet, ref_map) == [(2, normalized_reference)]

        sheet.cell(2, COUNT_COLUMN).value = "N"
        assert build_search_rows(sheet, ref_map) == []
        sheet.cell(2, COUNT_COLUMN).value = None
        sheet.cell(2, 5).value = None
        assert build_search_rows(sheet, ref_map) == []
    finally:
        workbook.close()


def test_lookup_rejects_blank_excel_reference_without_csv_fallback(tmp_path: Path) -> None:
    csv_path = tmp_path / "properties.csv"
    workbook_path = tmp_path / "ECM_Record.xlsx"
    _write_csv(csv_path, reference="CSV/ONLY")
    sheet_name = _write_workbook(workbook_path)
    workbook = load_workbook(workbook_path)
    try:
        sheet = workbook[sheet_name]
        sheet.cell(2, OUTGOING_REFERENCE_COLUMN).value = None
        with pytest.raises(LookupError, match="blank Excel outgoing reference"):
            build_search_rows(sheet, read_outgoing_reference_map(csv_path))
    finally:
        workbook.close()


@pytest.mark.parametrize(
    ("outgoing_reference", "revision", "expected"),
    (
        (
            "CWCRCC(ND202404)JV/1131/CSF/MS/0283-2025(F",
            "F",
            "CWCRCC(ND202404)JV/1131/CSF/MS/0283-2025",
        ),
        ("REF/42/F", "F", "REF/42"),
        ("REF/42(F)", "F", "REF/42"),
        ("REF/42(F", "G", "REF/42(F"),
        ("REF/42(F", "", "REF/42(F"),
    ),
)
def test_outgoing_reference_normalization_removes_only_matching_terminal_revision(
    outgoing_reference: str, revision: str, expected: str
) -> None:
    assert outgoing_reference_without_revision(outgoing_reference, revision) == expected


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
            self.value = value
            self.value = value

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


@pytest.mark.parametrize("operator_value", ["Starts With", "STARTSWITH"])
def test_lookup_accepts_visible_and_hidden_starts_with_operator_values(
    monkeypatch: pytest.MonkeyPatch, operator_value: str,
) -> None:
    import asyncio

    import control_panel_v3.workers.ecm_outgoing_reference_lookup as lookup

    class Control:
        def __init__(self, value: str = "", *, control_id: str = "") -> None:
            self.value = value
            self.control_id = control_id
            self.filled = None

        async def input_value(self):
            return self.value

        async def fill(self, value):
            self.filled = value
            self.value = value
            self.value = value

        async def dispatch_event(self, *_args):
            return None

        async def press(self, *_args):
            return None

    direction = Control("OUT")
    operator = Control(operator_value, control_id="OutgoingRefOperator")
    reference = Control()

    async def labeled(_form, label, *, include_operator=False, **_kwargs):
        if label == "In/Out":
            return [direction]
        return [operator, reference] if include_operator else [reference]

    async def metadata(control):
        return {"id": control.control_id, "class_name": ""}

    monkeypatch.setattr(lookup, "_labeled_controls", labeled)
    monkeypatch.setattr(lookup, "_control_metadata", metadata)
    asyncio.run(lookup._set_search_criteria(object(), "EXACT/REF"))
    assert reference.filled == "EXACT/REF"


def test_labeled_lookup_controls_include_operator_in_outgoing_reference_row(monkeypatch: pytest.MonkeyPatch) -> None:
    import asyncio

    import control_panel_v3.workers.ecm_step0_export as step0

    class Control:
        def __init__(self, control_id: str, class_name: str = "", *, control_type: str = "text", aria_hidden: str | None = None) -> None:
            self.control_id = control_id
            self.class_name = class_name
            self.control_type = control_type
            self.aria_hidden = aria_hidden

        async def is_visible(self):
            return True

        async def get_attribute(self, name):
            if name == "type":
                return self.control_type
            if name == "aria-hidden":
                return self.aria_hidden
            return None

    hover_help_button = Control("", "dijitOffScreen", control_type="button", aria_hidden="true")
    operator_arrow = Control("", "dijitReset dijitInputField dijitArrowButtonInner", aria_hidden="true")
    validation_icon = Control("", "dijitReset dijitInputField dijitValidationIcon dijitValidationInner")
    operator = Control("widget_SearchCriterian_4Operator", "dijitComboBox operatorSelect")
    reference = Control("ECM_OutRefNo", "dijitReset dijitInputInner")

    class Locator:
        def __init__(self, items):
            self.items = items

        async def count(self):
            return len(self.items)

        def nth(self, index):
            return self.items[index]

    class Row:
        def locator(self, _selector):
            return Locator([hover_help_button, operator, operator_arrow, validation_icon, reference])

    class Rows(Locator):
        def nth(self, index):
            assert index == 0
            return Row()

    class Form:
        def locator(self, selector):
            if selector.startswith("input,textarea"):
                return Locator([])
            if selector.startswith('tr:has([name="ECM_OutRefNo"])'):
                return Rows([Row()])
            raise AssertionError(f"Unexpected selector: {selector}")

    async def metadata(control):
        return {"widget": control.control_id or f"w-{id(control)}", "id": control.control_id, "class_name": control.class_name}

    monkeypatch.setattr(step0, "_control_metadata", metadata)
    controls = asyncio.run(step0._labeled_controls(Form(), "Outgoing ref. no.", include_operator=True))
    assert controls == [operator, reference]


@pytest.mark.parametrize(
    ("initial_value", "hidden_value", "should_select"),
    [("", "", True), ("", "2", False), ("OUT", "2", False)],
)
def test_lookup_sets_blank_in_out_and_accepts_saved_out_key(
    monkeypatch: pytest.MonkeyPatch, initial_value: str, hidden_value: str, should_select: bool,
) -> None:
    import asyncio

    import control_panel_v3.workers.ecm_outgoing_reference_lookup as lookup

    class Control:
        def __init__(self, value: str = "", *, control_id: str = "", class_name: str = "", hidden: str = "") -> None:
            self.value = value
            self.control_id = control_id
            self.class_name = class_name
            self.hidden = hidden
            self.filled = None
            self.pressed: list[str] = []

        async def input_value(self):
            return self.value

        async def evaluate(self, _script):
            return self.hidden

        async def fill(self, value):
            self.filled = value

        async def dispatch_event(self, event):
            if event == "change" and self.filled:
                self.value = self.filled
                self.hidden = "2"
            return None

        async def press(self, key):
            self.pressed.append(key)

    direction = Control(initial_value, hidden=hidden_value)
    operator = Control("STARTSWITH", control_id="", class_name="dijitComboBox operatorSelect")
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

    assert "Enter" not in direction.pressed
    assert (direction.filled == "OUT") is should_select
    assert direction.hidden == ("2" if should_select or hidden_value == "2" else "")
    assert reference.filled == "EXACT/REF"


def test_lookup_expands_collapsed_search_criteria_and_handles_empty_results_pane(monkeypatch: pytest.MonkeyPatch) -> None:
    import asyncio

    import control_panel_v3.workers.ecm_outgoing_reference_lookup as lookup
    import control_panel_v3.workers.ecm_step0_export as step0

    collapsed = [True]
    clicked_toggle = [False]

    class ToggleButton:
        async def is_visible(self):
            return True

    toggle_button = ToggleButton()

    class CountLocator:
        def __init__(self, count_fn, items=()):
            self._count_fn = count_fn
            self._items = items

        async def count(self):
            return self._count_fn()

        def nth(self, index):
            return self._items[index]

    class EmptyPane:
        async def is_visible(self):
            return True

        async def inner_text(self):
            return "No results were found."

    class CollapsiblePanel:
        def locator(self, selector: str):
            if selector == ".searchToggleArea.dijitHidden":
                return CountLocator(lambda: 1 if collapsed[0] else 0)
            if selector == '[data-dojo-attach-point="titleToggleNode"]':
                return CountLocator(lambda: 1, [toggle_button])
            if selector == '[data-dojo-attach-point="emptySearchResultsPane"]':
                return CountLocator(lambda: 1, [EmptyPane()])
            return CountLocator(lambda: 0)

        def get_by_role(self, *_args, **_kwargs):
            return CountLocator(lambda: 0)

    async def fake_click(target, _timeout_ms: int) -> None:
        assert target is toggle_button
        clicked_toggle[0] = True
        collapsed[0] = False

    monkeypatch.setattr(step0, "click_control", fake_click)
    panel = CollapsiblePanel()
    asyncio.run(step0._ensure_lookup_criteria_expanded(panel, 1000))
    assert clicked_toggle[0] is True
    assert collapsed[0] is False

    busy_checks = [0]

    class BusyThenIdlePage:
        async def evaluate(self, _script: str):
            busy_checks[0] += 1
            return busy_checks[0] == 1

    monkeypatch.setattr(lookup.asyncio, "sleep", lambda _seconds: asyncio.Future())
    # Use a real resolved coroutine for sleep in this isolated test.
    async def fast_sleep(_seconds: float) -> None:
        return None

    monkeypatch.setattr(lookup.asyncio, "sleep", fast_sleep)
    async def fresh_empty(*_args, **_kwargs):
        return False

    monkeypatch.setattr(lookup, "_wait_for_lookup_results", fresh_empty)
    assert asyncio.run(lookup._read_oldest_result_date(
        panel, BusyThenIdlePage(), 1000, before={}, query_id=1, row_number=2, outgoing_reference="REF/42"
    )) is None


def test_lookup_criteria_hydration_waits_for_required_properties(monkeypatch: pytest.MonkeyPatch) -> None:
    import asyncio

    import control_panel_v3.workers.ecm_step0_export as step0

    class Form:
        async def evaluate(self, script: str):
            assert "ECM_InOut" in script
            assert "ECM_OutRefNo" in script
            return {
                "input_area_present": True,
                "row_count": 2,
                "visible_row_count": 2,
                "missing_properties": [],
            }

    async def active_form(_panel):
        return Form()

    monkeypatch.setattr(step0, "_active_form", active_form)
    asyncio.run(step0._wait_for_lookup_criteria_hydration(object(), 100))


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
    async def fresh_results(_panel, _page, _timeout, _before):
        return True

    async def page_snapshot(_panel, _page):
        return {"busy": False, "mutations": 0, "text": "", "rows": [], "emptyText": ""}

    monkeypatch.setattr(lookup, "_wait_for_lookup_results", fresh_results)
    monkeypatch.setattr(lookup, "_lookup_results_snapshot", page_snapshot)
    panel, page = Panel(), Page()
    result = asyncio.run(lookup._read_oldest_result_date(
        panel, page, 1000, before={}, query_id=1, row_number=2, outgoing_reference="REF/42"
    ))
    assert result == date(2026, 8, 1)


def test_lookup_wait_rejects_unchanged_visible_grid(monkeypatch: pytest.MonkeyPatch) -> None:
    import asyncio

    import control_panel_v3.workers.ecm_outgoing_reference_lookup as lookup

    snapshot = {"busy": False, "text": "old result", "rows": ["old result"], "emptyText": ""}

    async def read_snapshot(_panel, _page):
        return snapshot

    async def fast_sleep(_seconds):
        return None

    monkeypatch.setattr(lookup, "_lookup_results_snapshot", read_snapshot)
    monkeypatch.setattr(lookup.asyncio, "sleep", fast_sleep)
    with pytest.raises(LookupError, match="fresh completed result"):
        asyncio.run(lookup._wait_for_lookup_results(object(), object(), 1, dict(snapshot)))


def test_lookup_wait_accepts_updated_grid_after_busy_transition(monkeypatch: pytest.MonkeyPatch) -> None:
    import asyncio

    import control_panel_v3.workers.ecm_outgoing_reference_lookup as lookup

    states = iter((
        {"busy": True, "text": "old", "rows": ["header", "old"], "emptyText": ""},
        {"busy": False, "text": "new", "rows": ["header", "new"], "emptyText": ""},
    ))

    async def read_snapshot(_panel, _page):
        return next(states)

    async def fast_sleep(_seconds):
        return None

    monkeypatch.setattr(lookup, "_lookup_results_snapshot", read_snapshot)
    monkeypatch.setattr(lookup.asyncio, "sleep", fast_sleep)
    assert asyncio.run(lookup._wait_for_lookup_results(
        object(), object(), 100, {"busy": False, "text": "old", "rows": ["header", "old"], "emptyText": ""}
    )) is True


@pytest.mark.parametrize("already_open,already_selected", [(False, False), (True, False), (False, True)])
@pytest.mark.parametrize("apply_succeeds", [True, False])
def test_lookup_opens_search_options_dialog_and_applies_version(monkeypatch: pytest.MonkeyPatch, already_open: bool, already_selected: bool, apply_succeeds: bool) -> None:
    import asyncio

    import control_panel_v3.workers.ecm_outgoing_reference_lookup as lookup

    state = {"open": already_open, "selected": already_selected, "menu": False, "applied": False, "clicks": []}

    class Collection:
        def __init__(self, items):
            self.items = items

    class Element:
        def __init__(self, name):
            self.name = name

        async def get_attribute(self, name):
            return "search_moreOptions_moreOptionsLink" if name == "id" else None

        async def evaluate(self, _script):
            return {"label": "All versions" if state["selected"] else "Released version",
                    "value": "allversions" if state["selected"] else "releasedversion"}

        async def inner_text(self):
            return "Search options: All versions" if state["applied"] else "Search options: Released version"

        def locator(self, selector_text):
            assert '_versionOptionNode' in selector_text
            return Collection([row] if state["open"] else [])

        def get_by_role(self, role, **_kwargs):
            if self.name == "dialog" and role == "button":
                return Collection([ok] if state["open"] else [])
            if self.name == "row" and role == "listbox":
                return Collection([selector])
            raise AssertionError((self.name, role))

    button, dialog, row, selector, option, ok = (Element(name) for name in ("button", "dialog", "row", "selector", "option", "ok"))

    class Page:
        def locator(self, selector_text):
            assert 'search_moreOptions_moreOptionsDialog' in selector_text
            return Collection(lambda: [dialog] if state["open"] else [])

        def get_by_role(self, role, **_kwargs):
            assert role == "option"
            return Collection([option] if state["menu"] else [])

    page = Page()
    button.page = page
    selector.page = page

    class Panel:
        def locator(self, selector_text):
            assert 'moreOptionsLink' in selector_text
            return Collection([button])

    async def visible(collection):
        return collection.items() if callable(collection.items) else collection.items

    async def click(target, _timeout):
        state["clicks"].append(target.name)
        if target is button:
            state["open"] = True
        elif target is selector:
            state["menu"] = True
        elif target is option:
            state["selected"] = True
            state["menu"] = False
        elif target is ok:
            state["applied"] = apply_succeeds
            state["open"] = False

    monkeypatch.setattr(lookup, "_visible_locators", visible)
    monkeypatch.setattr(lookup, "click_control", click)
    if apply_succeeds:
        asyncio.run(lookup._ensure_all_versions(Panel(), 100))
    else:
        with pytest.raises(LookupError, match="did not apply All versions"):
            asyncio.run(lookup._ensure_all_versions(Panel(), 100))
    assert state["applied"] is apply_succeeds
    assert ("button" in state["clicks"]) is not already_open
    assert ("option" in state["clicks"]) is not already_selected


def test_lookup_recovers_when_authentication_appears_during_saved_search(monkeypatch: pytest.MonkeyPatch) -> None:
    import asyncio

    import control_panel_v3.workers.ecm_outgoing_reference_lookup as lookup
    from control_panel_v3.workers.ecm_step0_export import AuthenticationRequired

    events = []

    async def auth_required(_page):
        return False

    async def open_search(_page, _timeout, **_kwargs):
        events.append("open")
        if events.count("open") == 1:
            raise AuthenticationRequired("login appeared")
        return "panel"

    async def wait_for_auth(_page, _predicate, _timeout):
        events.append("manual-auth")

    monkeypatch.setattr(lookup, "_page_requires_authentication", auth_required)
    monkeypatch.setattr(lookup, "_open_saved_search", open_search)
    monkeypatch.setattr(lookup, "wait_for_manual_auth", wait_for_auth)
    assert asyncio.run(lookup._open_lookup_with_auth(object(), 1)) == "panel"
    assert events == ["open", "manual-auth", "open"]


def test_lookup_late_authentication_timeout_does_not_retry_search(monkeypatch: pytest.MonkeyPatch) -> None:
    import asyncio

    import control_panel_v3.workers.ecm_outgoing_reference_lookup as lookup
    from control_panel_v3.workers.ecm_step0_export import AuthenticationRequired

    calls = []

    async def auth_required(_page):
        return False

    async def open_search(_page, _timeout, **_kwargs):
        calls.append("open")
        raise AuthenticationRequired("login appeared")

    async def wait_for_auth(_page, _predicate, _timeout):
        raise RuntimeError("authentication timed out")

    monkeypatch.setattr(lookup, "_page_requires_authentication", auth_required)
    monkeypatch.setattr(lookup, "_open_saved_search", open_search)
    monkeypatch.setattr(lookup, "wait_for_manual_auth", wait_for_auth)
    with pytest.raises(RuntimeError, match="authentication timed out"):
        asyncio.run(lookup._open_lookup_with_auth(object(), 1))
    assert calls == ["open"]


def test_lookup_detects_ecm_login_pane_when_text_probe_misses(monkeypatch: pytest.MonkeyPatch) -> None:
    import asyncio

    import control_panel_v3.workers.ecm_outgoing_reference_lookup as lookup

    class Page:
        def locator(self, selector):
            assert selector == '.ecmLoginPane [type="password"]'
            return object()

    async def no_auth(_page):
        return False

    async def visible(_locator):
        return [object()]

    monkeypatch.setattr(lookup, "_page_requires_authentication", no_auth)
    monkeypatch.setattr(lookup, "_visible_locators", visible)
    assert asyncio.run(lookup._lookup_requires_authentication(Page())) is True


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
        sheet.cell(row_number, OUTGOING_REFERENCE_COLUMN).value = "REF/HIT" if row_number == 2 else "REF/MISS"
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

    async def no_op(*_args, **_kwargs):
        return None

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
    async def no_authentication(_page):
        return False

    monkeypatch.setattr(lookup, "_lookup_requires_authentication", no_authentication)
    monkeypatch.setattr(lookup, "_ensure_all_versions", no_op)
    async def fresh_snapshot(*_args):
        return {"busy": False, "text": "before", "rows": [], "emptyText": ""}

    monkeypatch.setattr(lookup, "_lookup_results_snapshot", fresh_snapshot)
    monkeypatch.setattr(lookup, "_active_form", fake_active_form)
    monkeypatch.setattr(lookup, "_set_search_criteria", fake_set_criteria)
    monkeypatch.setattr(lookup, "_visible_locators", fake_visible)
    async def fake_read(panel, page, timeout, **_kwargs):
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

    authentication_required = [True]
    focus_events: list[str] = []

    class Field:
        async def is_visible(self):
            return True

        async def is_enabled(self):
            return True

        async def click(self, **_kwargs):
            focus_events.append("username")

    class Fields:
        async def count(self):
            return 1

        def nth(self, _index):
            return Field()

    class Page:
        frames = []

        async def bring_to_front(self):
            focus_events.append("front")

        def locator(self, _selector):
            return Fields()

    async def fake_sleep(_value):
        if focus_events.count("poll") == 0:
            focus_events.append("poll")
        else:
            authentication_required[0] = False

    async def requires_authentication(_page):
        return authentication_required[0]

    monkeypatch.setattr("control_panel_v3.workers.browser_support.asyncio.sleep", fake_sleep)
    asyncio.run(wait_for_manual_auth(Page(), requires_authentication, 1))
    assert not authentication_required[0]
    assert focus_events.count("front") == 1
    assert focus_events.count("username") == 1


def test_browser_support_login_focus_is_best_effort_without_username_field() -> None:
    import asyncio

    from control_panel_v3.workers.browser_support import focus_login_username

    class Page:
        frames = []

        async def bring_to_front(self):
            return None

        def locator(self, _selector):
            raise RuntimeError("No login form")

    asyncio.run(focus_login_username(Page()))


def test_browser_support_authentication_timeout_and_closed_page() -> None:
    import asyncio

    from control_panel_v3.workers.browser_support import wait_for_manual_auth

    async def requires_authentication(_page):
        return True

    class ClosedPage:
        def is_closed(self):
            return True

    with pytest.raises(RuntimeError, match="authentication timed out"):
        asyncio.run(wait_for_manual_auth(object(), requires_authentication, 0.01))
    with pytest.raises(RuntimeError, match="closed during authentication"):
        asyncio.run(wait_for_manual_auth(ClosedPage(), requires_authentication, 1))


@pytest.mark.parametrize("scenario", ["signed-in", "existing-session", "auth-timeout", "search-failure"])
def test_lookup_authentication_handoff_and_failure_diagnostics(tmp_path: Path, monkeypatch: pytest.MonkeyPatch, scenario: str) -> None:
    import asyncio
    import types

    import control_panel_v3.workers.ecm_outgoing_reference_lookup as lookup

    # Make the worker's V3-only path boundary point to this isolated fixture.
    monkeypatch.setattr(lookup, "__file__", str(tmp_path / "workers" / "lookup.py"))
    csv_path = tmp_path / "properties.csv"
    workbook_path = tmp_path / "review.xlsx"
    _write_csv(csv_path)
    sheet_name = _write_workbook(workbook_path)
    original = workbook_path.read_bytes()
    events = []

    class Page:
        async def goto(self, *_args, **_kwargs):
            events.append("goto")

        def is_closed(self):
            return False

    page = Page()

    class Context:
        pages = [page]

        async def close(self):
            events.append("close")

    class Chromium:
        async def launch_persistent_context(self, **kwargs):
            assert kwargs["headless"] is False
            assert kwargs["user_data_dir"] == str(tmp_path / "profile")
            return Context()

    class Playwright:
        chromium = Chromium()

        async def __aenter__(self):
            return self

        async def __aexit__(self, *_args):
            return None

    monkeypatch.setitem(sys.modules, "playwright.async_api", types.SimpleNamespace(async_playwright=Playwright))
    auth_checks = [0]
    auth_state = [True]

    async def requires_authentication(current_page):
        assert current_page is page
        auth_checks[0] += 1
        if scenario == "existing-session":
            auth_state[0] = False
        elif scenario == "signed-in":
            auth_state[0] = auth_checks[0] <= 2
        else:
            auth_state[0] = True if scenario == "auth-timeout" else auth_checks[0] == 1
        return auth_state[0]

    async def open_saved_search(current_page, _timeout, **_kwargs):
        assert current_page is page
        assert not auth_state[0], "Search must not open while authentication is required"
        events.append("open-search")
        if scenario == "search-failure":
            raise RuntimeError("search unavailable")
        return object()

    async def diagnostics(current_page, _directory, _stage, _diagnostics):
        assert current_page is page
        assert "close" not in events
        events.append("diagnostics")

    async def return_object(*_args, **_kwargs):
        return object()

    async def return_buttons(*_args, **_kwargs):
        return [object()]

    async def no_op(*_args, **_kwargs):
        return None

    async def first_date(*_args, **_kwargs):
        return date(2026, 8, 1)

    monkeypatch.setattr(lookup, "_page_requires_authentication", requires_authentication)
    monkeypatch.setattr(lookup, "save_failure_diagnostics", diagnostics)
    monkeypatch.setattr(lookup, "_active_form", return_object)
    monkeypatch.setattr(lookup, "_set_search_criteria", no_op)
    monkeypatch.setattr(lookup, "_visible_locators", return_buttons)
    monkeypatch.setattr(lookup, "click_control", no_op)
    monkeypatch.setattr(lookup, "_read_oldest_result_date", first_date)

    # The fake panel must expose the Search button lookup used by the worker.
    class Panel:
        def get_by_role(self, *_args, **_kwargs):
            return object()

    panel = Panel()

    async def open_panel(current_page, timeout, **kwargs):
        await open_saved_search(current_page, timeout, **kwargs)
        return panel

    monkeypatch.setattr(lookup, "_open_saved_search", open_panel)
    monkeypatch.setattr(lookup, "_ensure_all_versions", no_op)
    async def fresh_snapshot(*_args):
        return {"busy": False, "text": "before", "rows": [], "emptyText": ""}

    monkeypatch.setattr(lookup, "_lookup_results_snapshot", fresh_snapshot)
    params = dict(workbook_path=workbook_path, worksheet_name=sheet_name, csv_path=csv_path,
                  profile_dir=tmp_path / "profile", ecm_url="https://example.test/", timeout_seconds=0.01,
                  summary_path=tmp_path / "summary.json", diagnostics_dir=tmp_path / "diagnostics")
    if scenario in {"signed-in", "existing-session"}:
        if scenario == "signed-in":
            async def no_delay(_seconds):
                return None
            monkeypatch.setattr("control_panel_v3.workers.browser_support.asyncio.sleep", no_delay)
        assert asyncio.run(lookup.run_lookup(**params)) == (1, 0, 1)
        updated = load_workbook(workbook_path, read_only=True)
        try:
            assert updated[sheet_name].cell(2, REMARK_COLUMN).value == "2026-08-01"
        finally:
            updated.close()
        assert "diagnostics" not in events
    else:
        message = "authentication timed out" if scenario == "auth-timeout" else "saved search"
        with pytest.raises(RuntimeError, match=message):
            asyncio.run(lookup.run_lookup(**params))
        assert workbook_path.read_bytes() == original
        assert not params["summary_path"].exists()
        assert events.index("diagnostics") < events.index("close")
        if scenario == "auth-timeout":
            assert "open-search" not in events


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
        step1a_run = history.list_runs(workflow_id="ecm-step-1a", period_id=period.id)[0]
        assert accepted.headers["location"] == f"/runs/{step1a_run.id}"
        detail = client.get(accepted.headers["location"])
        assert detail.status_code == 200
        assert f"RUN #{step1a_run.id}" in detail.text
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
        assert lookup_response.headers["location"] == "/report-history"
        lookup_run = history.list_runs(workflow_id="ecm-first-date-lookup", period_id=lookup_period.id)[0]
        assert lookup_run.lookup_scope == "blank"

        bypass_period = history.create_report_period("2026-08-15", "2026-08-21")
        bypass_response = client.post(
            f"/report-periods/{bypass_period.id}/runs/ecm-step-2",
            data={"confirm_step_1b_bypass": "true"}, follow_redirects=False,
        )
        assert bypass_response.status_code == 303
        assert bypass_response.headers["location"] == "/report-history"
        bypass_run = history.list_runs(workflow_id="ecm-step-2", period_id=bypass_period.id)[0]
        assert bypass_run.manual_step_1b_bypass is True
        assert "Manual Step 1b bypass confirmed" in client.get("/history").text
    finally:
        client.close()
        import shutil
        shutil.rmtree(test_root, ignore_errors=True)


@pytest.mark.parametrize("recent_has_search", [True, False])
def test_lookup_open_search_waits_for_panel_and_avoids_cross_group_ambiguity(monkeypatch: pytest.MonkeyPatch, recent_has_search: bool) -> None:
    import asyncio
    from control_panel_v3.workers import ecm_step0_export as step0

    class Locator:
        def __init__(self, name: str) -> None:
            self.name = name
            self.first = self

        async def wait_for(self, **_kwargs) -> None:
            assert opened[0]

    class Page:
        def get_by_role(self, _role: str, name: str, **_kwargs) -> Locator:
            return Locator(name)

    opened = [False]
    polls = [0]
    events: list[str] = []

    async def wait_for_search_panel(_page, _timeout_ms: int) -> object:
        events.append("wait_for_search_panel")
        return object()

    async def visible(locator):
        if locator.name == "Search ECM Document":
            return [locator] if opened[0] else []
        if locator.name in {"Recent Searches", "All Searches"}:
            polls[0] += 1
            # Simulate delayed Dojo tree rendering after authentication clears.
            if polls[0] < 3:
                return []
            return [locator]
        return []

    async def expand(locator, _timeout_ms: int) -> None:
        events.append(f"expand:{locator.name}")

    async def saved(_page, group: str, search_name: str):
        assert search_name == "Search ECM Document"
        # Present in All Searches always, and also in Recent Searches when recent_has_search=True.
        if group == "All Searches" or recent_has_search:
            return [Locator(f"{group}:{search_name}")]
        return []

    async def click(locator, _timeout_ms: int) -> None:
        events.append(f"click:{locator.name}")
        opened[0] = True

    async def no_delay(_seconds: float) -> None:
        return None

    async def active_form(panel):
        return panel

    monkeypatch.setattr(step0, "_wait_for_search_panel", wait_for_search_panel)
    monkeypatch.setattr(step0, "_visible_locators", visible)
    monkeypatch.setattr(step0, "_expand_treeitem", expand)
    monkeypatch.setattr(step0, "_saved_searches_in_group", saved)
    monkeypatch.setattr(step0, "click_control", click)
    monkeypatch.setattr(step0, "_active_form", active_form)
    monkeypatch.setattr(step0.asyncio, "sleep", no_delay)

    panel = asyncio.run(step0._open_lookup_search(Page(), 5000, search_name="Search ECM Document"))
    assert panel.name == "Search ECM Document"
    expected_prefix = ["wait_for_search_panel", "expand:Recent Searches", "click:Recent Searches:Search ECM Document"]
    if not recent_has_search:
        expected_prefix = [
            "wait_for_search_panel",
            "expand:Recent Searches",
            "expand:All Searches",
            "click:All Searches:Search ECM Document",
        ]
    assert events == expected_prefix
