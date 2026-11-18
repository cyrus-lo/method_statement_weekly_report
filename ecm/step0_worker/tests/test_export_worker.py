from __future__ import annotations

import csv
import sys
import types
from datetime import date
from pathlib import Path

import pytest

WORKSPACE_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(WORKSPACE_ROOT))

from ecm.step0_worker.export_worker import (
    BrowserDiagnostics,
    CSV_HEADERS,
    CsvValidationError,
    NoSearchResults,
    PageNotReadyError,
    Step0Error,
    CURSOR_PAUSE_MAX_SECONDS,
    CURSOR_TRAVEL_MAX_SECONDS,
    CURSOR_TRAVEL_MIN_SECONDS,
    _PageCursor,
    _save_failure_diagnostics,
    _reset_profile,
    _click_ui_control,
    _wait_for_ecm_readiness,
    _export_csv,
    _date_input_value,
    validate_properties_csv,
)
from control_panel.services.workflows import _step0_output_path


def _write_csv(path: Path, rows: list[dict[str, str]], headers: tuple[str, ...] = CSV_HEADERS) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as source:
        writer = csv.DictWriter(source, fieldnames=headers, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def _valid_row() -> dict[str, str]:
    return {
        "Document title": "Method Statement",
        "Document date": "7/20/2026, 12:00 AM",
        "Name": "example.pdf",
        "Document revision": "A",
        "Modified By": "dm_admin",
        "Major Version": "1",
        "SMM ID": "123",
        "Modified On": "7/20/2026, 2:18 PM",
        "Outgoing ref. no.": "example",
        "Path": "/example.pdf",
        "Link": "https://example.test/example.pdf",
    }


def test_validate_properties_csv_accepts_bom_csv(tmp_path: Path) -> None:
    csv_path = tmp_path / "properties.csv"
    _write_csv(csv_path, [_valid_row()])

    assert validate_properties_csv(csv_path) == 1


def test_validate_properties_csv_allows_optional_blank_values(tmp_path: Path) -> None:
    csv_path = tmp_path / "properties.csv"
    row = _valid_row()
    row["Document revision"] = ""
    row["SMM ID"] = ""
    row["Modified On"] = ""
    _write_csv(csv_path, [row])

    assert validate_properties_csv(csv_path) == 1


def test_validate_properties_csv_rejects_missing_file(tmp_path: Path) -> None:
    with pytest.raises(CsvValidationError, match="not found"):
        validate_properties_csv(tmp_path / "missing.csv")


def test_validate_properties_csv_rejects_missing_header(tmp_path: Path) -> None:
    csv_path = tmp_path / "properties.csv"
    headers = tuple(header for header in CSV_HEADERS if header != "SMM ID")
    _write_csv(csv_path, [_valid_row()], headers)

    with pytest.raises(CsvValidationError, match="SMM ID"):
        validate_properties_csv(csv_path)


def test_validate_properties_csv_rejects_invalid_date(tmp_path: Path) -> None:
    csv_path = tmp_path / "properties.csv"
    row = _valid_row()
    row["Document date"] = "2026-07-20"
    _write_csv(csv_path, [row])

    with pytest.raises(CsvValidationError, match="Document date"):
        validate_properties_csv(csv_path)


def test_validate_properties_csv_rejects_invalid_major_version(tmp_path: Path) -> None:
    csv_path = tmp_path / "properties.csv"
    row = _valid_row()
    row["Major Version"] = "not-an-integer"
    _write_csv(csv_path, [row])

    with pytest.raises(CsvValidationError, match="Major Version"):
        validate_properties_csv(csv_path)


def test_validate_properties_csv_rejects_blank_required_field(tmp_path: Path) -> None:
    csv_path = tmp_path / "properties.csv"
    row = _valid_row()
    row["Name"] = ""
    _write_csv(csv_path, [row])

    with pytest.raises(CsvValidationError, match="Name"):
        validate_properties_csv(csv_path)


def test_validate_properties_csv_rejects_empty_export(tmp_path: Path) -> None:
    csv_path = tmp_path / "properties.csv"
    _write_csv(csv_path, [])

    with pytest.raises(NoSearchResults, match="no document rows"):
        validate_properties_csv(csv_path)


def test_step0_output_path_is_timestamped_and_period_specific(tmp_path: Path) -> None:
    output = _step0_output_path(tmp_path, date(2026, 7, 20), date(2026, 7, 26))

    assert output.parent == tmp_path
    assert output.name.startswith("properties_2026-07-20_to_2026-07-26_")
    assert output.suffix == ".csv"


def test_date_input_value_uses_ecm_short_date_format() -> None:
    assert _date_input_value(date(2026, 7, 27)) == "7/27/2026"


def test_page_cursor_duration_is_distance_sensitive_and_bounded() -> None:
    class FakePage:
        pass

    cursor = _PageCursor(FakePage())

    short = cursor._travel_duration(10)
    long = cursor._travel_duration(900)

    assert CURSOR_TRAVEL_MIN_SECONDS <= short < long <= CURSOR_TRAVEL_MAX_SECONDS


def test_page_cursor_moves_in_waypoints_and_caps_optional_pauses(monkeypatch) -> None:
    import asyncio
    import random

    moves = []
    sleeps = []

    class FakeMouse:
        async def move(self, x, y):
            moves.append((x, y))

    class FakePage:
        mouse = FakeMouse()

    async def fake_sleep(seconds):
        sleeps.append(seconds)

    monkeypatch.setattr(asyncio, "sleep", fake_sleep)
    cursor = _PageCursor(FakePage(), rng=random.Random(7))
    asyncio.run(cursor.move_to_box({"x": 500, "y": 200, "width": 100, "height": 40}))

    assert len(moves) >= 8
    assert 500 <= moves[-1][0] <= 600
    assert 200 <= moves[-1][1] <= 240
    assert all(0 < pause <= CURSOR_PAUSE_MAX_SECONDS for pause in sleeps)


def test_click_ui_control_moves_cursor_before_locator_click(monkeypatch) -> None:
    import asyncio

    events = []

    class FakeMouse:
        async def move(self, x, y):
            events.append("move")

    class FakePage:
        mouse = FakeMouse()

    class FakeLocator:
        page = FakePage()

        async def bounding_box(self):
            return {"x": 10, "y": 20, "width": 100, "height": 30}

        async def click(self, **kwargs):
            events.append("click")

    async def fake_sleep(seconds):
        return None

    monkeypatch.setattr(asyncio, "sleep", fake_sleep)
    asyncio.run(_click_ui_control(FakeLocator(), 1000))

    assert events[-1] == "click"
    assert "move" in events[:-1]


def test_browser_diagnostics_keeps_only_bounded_safe_events() -> None:
    callbacks = {}

    class FakePage:
        def on(self, event, callback):
            callbacks[event] = callback

    class FakeMessage:
        type = "error"
        text = "x" * 600
        location = {"url": "https://example.test/app?token=hidden"}

    diagnostics = BrowserDiagnostics(limit=2)
    diagnostics.attach(FakePage())
    callbacks["console"](FakeMessage())
    callbacks["console"](FakeMessage())
    callbacks["console"](FakeMessage())

    assert len(diagnostics.console) == 2
    assert diagnostics.console[0]["text"] == "x" * 500


def test_reset_profile_moves_existing_profile_to_timestamped_backup(tmp_path: Path) -> None:
    profile_dir = tmp_path / ".step0-edge-profile"
    profile_dir.mkdir()
    (profile_dir / "Cookies").write_text("preserve", encoding="utf-8")

    backup_dir = _reset_profile(profile_dir)

    assert backup_dir is not None
    assert not profile_dir.exists()
    assert backup_dir.name.startswith(".step0-edge-profile.backup-")
    assert (backup_dir / "Cookies").read_text(encoding="utf-8") == "preserve"


def test_browser_diagnostics_logs_safe_browser_events(monkeypatch) -> None:
    callbacks = {}
    logged: list[tuple[str, dict[str, object]]] = []

    class FakePage:
        def on(self, event, callback):
            callbacks[event] = callback

    class FakeMessage:
        type = "warning"
        text = "warning text"
        location = {
            "url": "https://example.test/app?token=hidden#fragment",
            "lineNumber": 12,
            "columnNumber": 4,
        }

    class FakeRequest:
        method = "GET"
        resource_type = "script"
        url = "https://example.test/app.js?token=hidden"
        failure = "blocked by client"

    class FakeResponseRequest:
        method = "GET"

    class FakeResponse:
        status = 503
        url = "https://example.test/api?token=hidden"
        request = FakeResponseRequest()

    monkeypatch.setattr(
        "ecm.step0_worker.export_worker._log",
        lambda event, **fields: logged.append((event, fields)),
    )
    diagnostics = BrowserDiagnostics(limit=2)
    diagnostics.attach(FakePage())
    callbacks["console"](FakeMessage())
    callbacks["pageerror"](RuntimeError("script failed"))
    callbacks["requestfailed"](FakeRequest())
    callbacks["response"](FakeResponse())

    assert [event for event, _ in logged] == [
        "browser_console",
        "browser_page_error",
        "browser_request_failed",
        "browser_response_failed",
    ]
    assert logged[0][1]["location"] == {
        "url": "https://example.test/app",
        "lineNumber": 12,
        "columnNumber": 4,
    }
    assert "token=hidden" not in str(logged)
    assert diagnostics.request_failures[0]["error"] == "blocked by client"
    assert diagnostics.response_failures[0]["status"] == 503


def test_browser_diagnostics_filters_and_survives_logging_errors(monkeypatch) -> None:
    callbacks = {}

    class FakePage:
        def on(self, event, callback):
            callbacks[event] = callback

    class NormalConsole:
        type = "log"
        text = "ignored"
        location = {}

    class BrokenEvent:
        @property
        def status(self):
            raise RuntimeError("incomplete fake")

    def fail_log(*args, **kwargs):
        raise RuntimeError("log unavailable")

    monkeypatch.setattr("ecm.step0_worker.export_worker._log", fail_log)
    diagnostics = BrowserDiagnostics()
    diagnostics.attach(FakePage())
    callbacks["console"](NormalConsole())
    callbacks["response"](BrokenEvent())
    callbacks["pageerror"](RuntimeError("still safe"))

    assert diagnostics.console == []
    assert diagnostics.response_failures == []
    assert diagnostics.page_errors == [{"message": "still safe"}]


def test_save_failure_diagnostics_writes_metadata_and_visual_evidence(tmp_path: Path) -> None:
    import asyncio

    class FakePage:
        url = "https://example.test/navigator/?token=hidden"

        async def screenshot(self, path, full_page):
            Path(path).write_bytes(b"png")

        async def content(self):
            return "<html><body>safe test document</body></html>"

        async def title(self):
            return "Navigator"

    diagnostics = BrowserDiagnostics()
    diagnostics.page_errors.append({"message": "test"})
    asyncio.run(
        _save_failure_diagnostics(
            FakePage(), tmp_path, "waiting for ECM Navigator readiness",
            browser_diagnostics=diagnostics,
            readiness_state={"ready_state": "complete", "body_child_count": 0},
        )
    )

    metadata = next(tmp_path.glob("*.diagnostics.json"))
    content = metadata.read_text(encoding="utf-8")
    assert "token=hidden" not in content
    assert "Navigator" in content
    assert next(tmp_path.glob("*.png")).read_bytes() == b"png"
    assert next(tmp_path.glob("*.html")).is_file()


def test_page_readiness_rejects_empty_dom_before_search() -> None:
    import asyncio

    class FakeLocator:
        async def count(self):
            return 0

        async def inner_text(self):
            return ""

    class FakePage:
        url = "https://example.test/navigator/"

        def locator(self, selector):
            return FakeLocator()

        async def title(self):
            return ""

        async def evaluate(self, script):
            return {
                "ready_state": "complete",
                "body_text_length": 0,
                "body_html_length": 0,
                "body_child_count": 0,
                "iframe_count": 0,
            }

        async def reload(self, **kwargs):
            return None

    class FakeContext:
        pages = [FakePage()]

    with pytest.raises(PageNotReadyError, match="empty_dom"):
        asyncio.run(
            _wait_for_ecm_readiness(FakePage(), "https://example.test/navigator/", FakeContext(), 1)
        )


def test_export_all_confirmation_dialog_downloads_once(tmp_path: Path, monkeypatch) -> None:
    import asyncio

    from ecm.step0_worker import export_worker

    clicked: list[str] = []

    class FakeLocator:
        def __init__(self, name: str, checked: bool = False):
            self.name = name
            self.checked = checked

        @property
        def first(self):
            return self

        async def wait_for(self, **kwargs):
            return None

        async def is_checked(self):
            return self.checked

        def get_by_role(self, role, name=None, exact=False):
            return self.children[(role, name)]

    export_all = FakeLocator("Export All")
    dialog = FakeLocator("Export All dialog")
    all_items = FakeLocator("Export all items", checked=True)
    raw_values = FakeLocator("Export raw values", checked=False)
    export_action = FakeLocator("Export")
    dialog.children = {
        ("radio", "Export all items"): all_items,
        ("checkbox", "Export raw values"): raw_values,
        ("button", "Export"): export_action,
    }

    class FakePanel:
        def get_by_role(self, role, name=None, exact=False):
            return export_all

    class FakeDownload:
        save_count = 0

        async def failure(self):
            return None

        async def save_as(self, path):
            self.save_count += 1
            Path(path).write_bytes(b"downloaded CSV")

    class FakeDownloadContext:
        enter_count = 0

        def __init__(self):
            self.download = FakeDownload()
            self.value = self._download_value()

        async def _download_value(self):
            return self.download

        async def __aenter__(self):
            self.enter_count += 1
            return self

        async def __aexit__(self, *args):
            return None

    download_context = FakeDownloadContext()

    class FakePage:
        def get_by_role(self, role, name=None, exact=False):
            if role == "dialog":
                return dialog
            raise AssertionError(f"Unexpected page role lookup: {role} {name}")

        def expect_download(self, timeout):
            assert timeout == 500
            return download_context

    async def fake_visible(locator):
        return [locator]

    async def fake_click(locator, timeout_ms):
        clicked.append(locator.name)

    monkeypatch.setattr(export_worker, "_visible_locators", fake_visible)
    monkeypatch.setattr(export_worker, "_click_ui_control", fake_click)

    output = tmp_path / "properties.csv"
    asyncio.run(_export_csv(FakePage(), FakePanel(), output, 500))

    assert clicked == ["Export All", "Export"]
    assert download_context.enter_count == 1
    assert download_context.download.save_count == 1
    assert output.read_bytes() == b"downloaded CSV"


def test_saved_search_prefers_recent_searches(monkeypatch) -> None:
    import asyncio

    from ecm.step0_worker import export_worker

    clicked = []

    class FakeLocator:
        def __init__(self, name, visible=True, expanded=None):
            self.name = name
            self.visible = visible
            self.expanded = expanded

        async def get_attribute(self, name):
            return self.expanded if name == "aria-expanded" else None

        async def wait_for(self, **kwargs):
            return None

    recent = FakeLocator("Recent Searches", expanded="true")
    all_searches = FakeLocator("All Searches", visible=False, expanded="false")
    recent_group = FakeLocator("Recent Searches group")
    saved = FakeLocator("SMM_ECM_Reporting")

    class FakePage:
        def get_by_role(self, role, name=None, exact=False):
            if name == "Search":
                return FakeLocator("Search panel")
            if role == "group" and name == "Recent Searches":
                return recent_group
            if name == "Recent Searches":
                return recent
            if name == "All Searches":
                return all_searches
            if name == "SMM_ECM_Reporting":
                return saved
            return FakeLocator(str(name))

    def recent_group_visible_searches(*args, **kwargs):
        return saved

    recent_group.get_by_role = recent_group_visible_searches

    async def fake_wait(page, timeout_ms):
        return FakeLocator("Search panel")

    async def fake_visible(locator):
        if locator is saved or locator is recent_group:
            return [saved] if locator is saved else [recent_group]
        return [recent] if locator is recent else []

    async def fake_click(locator, timeout_ms):
        clicked.append(locator.name)

    async def fake_panel_wait(self, **kwargs):
        return None

    hydration_calls = []

    async def fake_wait_for_criteria_hydration(panel, timeout_ms):
        hydration_calls.append((panel, timeout_ms))

    monkeypatch.setattr(export_worker, "_wait_for_search_panel", fake_wait)
    monkeypatch.setattr(export_worker, "_visible_locators", fake_visible)
    monkeypatch.setattr(export_worker, "_click_ui_control", fake_click)
    monkeypatch.setattr(
        export_worker,
        "_wait_for_criteria_hydration",
        fake_wait_for_criteria_hydration,
    )
    monkeypatch.setattr(FakeLocator, "wait_for", fake_panel_wait)

    assert asyncio.run(export_worker._open_saved_search(FakePage(), 1000))
    assert clicked == ["SMM_ECM_Reporting"]
    assert len(hydration_calls) == 1
    assert hydration_calls[0][1] == 1000


def test_criteria_hydration_waits_for_delayed_dojo_rows(monkeypatch) -> None:
    import asyncio

    from ecm.step0_worker import export_worker

    states = iter(
        [
            {
                "probe_supported": True,
                "input_area_present": True,
                "row_count": 0,
                "visible_row_count": 0,
                "expected_property_count": 0,
                "missing_properties": [
                    "ECM_DocDate",
                    "ECM_OCRStatus",
                    "ECM_Title",
                    "ECM_InOut",
                ],
                "search_button_disabled": True,
            },
            {
                "probe_supported": True,
                "input_area_present": True,
                "row_count": 4,
                "visible_row_count": 4,
                "expected_property_count": 4,
                "missing_properties": [],
                "search_button_disabled": False,
            },
        ]
    )
    evaluations = []

    class FakeCriteriaForm:
        async def evaluate(self, script):
            evaluations.append(script)
            return next(states)

    events = []
    monkeypatch.setattr(
        export_worker,
        "_log",
        lambda event, **fields: events.append((event, fields)),
    )

    asyncio.run(export_worker._wait_for_criteria_hydration(FakeCriteriaForm(), 1000))

    assert len(evaluations) == 2
    assert [event for event, _ in events] == [
        "criteria_hydration_wait_started",
        "criteria_hydration_state",
        "criteria_hydration_state",
        "criteria_hydration_ready",
    ]


def test_criteria_hydration_timeout_is_distinct(monkeypatch) -> None:
    import asyncio

    from ecm.step0_worker import export_worker

    state = {
        "probe_supported": True,
        "input_area_present": True,
        "row_count": 0,
        "visible_row_count": 0,
        "expected_property_count": 0,
        "missing_properties": ["ECM_OCRStatus"],
        "search_button_disabled": True,
    }

    class FakeCriteriaForm:
        async def evaluate(self, script):
            return state

    with pytest.raises(export_worker.CriteriaHydrationError) as error:
        asyncio.run(export_worker._wait_for_criteria_hydration(FakeCriteriaForm(), 1))

    assert error.value.state == state


def test_dojo_operator_value_uses_hidden_value_when_visible_textbox_is_blank() -> None:
    import asyncio

    from ecm.step0_worker import export_worker

    class FakeOperator:
        async def input_value(self):
            return ""

        async def evaluate(self, script):
            return "BETWEEN"

        async def inner_text(self):
            return ""

    assert asyncio.run(export_worker._control_value(FakeOperator())) == "BETWEEN"


def test_run_export_refuses_existing_output_before_browser_start(tmp_path: Path) -> None:
    from ecm.step0_worker.export_worker import run_export

    output = tmp_path / "properties.csv"
    output.write_text("existing", encoding="utf-8")

    with pytest.raises(Step0Error, match="Refusing to overwrite"):
        import asyncio

        asyncio.run(
            run_export(
                start_date=date(2026, 7, 20),
                end_date=date(2026, 7, 26),
                output_path=output,
                ecm_url="https://example.test/",
                profile_dir=tmp_path / "profile",
                timeout_seconds=1,
            )
        )


def test_run_export_replaces_existing_period_input_after_validation(
    tmp_path: Path, monkeypatch
) -> None:
    from ecm.step0_worker.export_worker import run_export

    output = tmp_path / "properties-timestamped.csv"
    period_input = tmp_path / "properties-period.csv"
    period_input.write_text("old handoff", encoding="utf-8")

    valid_csv = tmp_path / "download.csv"
    _write_csv(valid_csv, [_valid_row()])

    async def fake_export_csv(page, search_panel, temporary_path, timeout_ms):
        temporary_path.write_bytes(valid_csv.read_bytes())

    monkeypatch.setattr(
        "ecm.step0_worker.export_worker._export_csv", fake_export_csv
    )

    class FakePage:
        async def goto(self, *args, **kwargs):
            return None

    class FakeContext:
        pages = [FakePage()]

        async def close(self):
            return None

    class FakeChromium:
        async def launch_persistent_context(self, **kwargs):
            return FakeContext()

    class FakePlaywright:
        chromium = FakeChromium()

    class FakePlaywrightContext:
        async def __aenter__(self):
            return FakePlaywright()

        async def __aexit__(self, *args):
            return None

    class FakeAsyncPlaywright:
        def __call__(self):
            return FakePlaywrightContext()

    async def fake_open_saved_search(page, timeout_ms):
        return object()

    async def fake_search_period(page, search_panel, start_date, end_date, timeout_ms):
        return 1

    async_api = types.ModuleType("playwright.async_api")
    async_api.async_playwright = FakeAsyncPlaywright()
    playwright = types.ModuleType("playwright")
    playwright.async_api = async_api
    monkeypatch.setitem(sys.modules, "playwright", playwright)
    monkeypatch.setitem(sys.modules, "playwright.async_api", async_api)
    monkeypatch.setattr(
        "ecm.step0_worker.export_worker._open_saved_search", fake_open_saved_search
    )
    async def fake_wait_for_ecm_readiness(page, ecm_url, context, timeout_ms):
        return {"ready_state": "complete"}

    monkeypatch.setattr(
        "ecm.step0_worker.export_worker._wait_for_ecm_readiness", fake_wait_for_ecm_readiness
    )
    monkeypatch.setattr(
        "ecm.step0_worker.export_worker._search_period", fake_search_period
    )

    import asyncio

    assert asyncio.run(
        run_export(
            start_date=date(2026, 7, 20),
            end_date=date(2026, 7, 26),
            output_path=output,
            period_input_path=period_input,
            ecm_url="https://example.test/",
            profile_dir=tmp_path / "profile",
            timeout_seconds=1,
        )
    ) == 1
    assert validate_properties_csv(output) == 1
    assert validate_properties_csv(period_input) == 1