from __future__ import annotations

import asyncio
import sys
import types
from datetime import date

import pytest

from control_panel_v3.workers import ecm_step0_export as worker


def test_cli_defaults_and_options():
    base = ["--start-date", "2026-01-01", "--end-date", "2026-01-07", "--output", "x", "--profile-dir", "profile"]
    args = worker.build_parser().parse_args(base)
    assert (args.headed, args.login_if_needed, args.timeout_seconds, args.auth_timeout_seconds, args.failure_browser_hold_seconds) == (True, True, 60, 120, 120)
    args = worker.build_parser().parse_args(base + ["--no-headed", "--no-login-if-needed", "--failure-browser-hold-seconds", "0"])
    assert not args.headed and not args.login_if_needed and args.failure_browser_hold_seconds == 0


def test_select_existing_ecm_page():
    first = types.SimpleNamespace(url="about:blank")
    target = types.SimpleNamespace(url="https://ecm.example/navigator/?secret=1")
    context = types.SimpleNamespace(pages=[first, target])
    assert asyncio.run(worker._select_ecm_page(context, "https://ecm.example/navigator/")) is target


def test_readiness_distinguishes_wrong_origin_and_auth(monkeypatch):
    page = types.SimpleNamespace(is_closed=lambda: False)
    context = object()

    async def state(*_args):
        return {"origin_matches": False, "ready_state": "complete", "body_child_count": 1, "body_text_length": 4, "shell_counts": {"main": 1}}

    async def no_auth(_page):
        return False

    monkeypatch.setattr(worker, "_ecm_page_state", state)
    monkeypatch.setattr(worker, "_page_requires_authentication", no_auth)
    with pytest.raises(worker.PageNotReadyError) as error:
        asyncio.run(worker._wait_for_ecm_readiness(page, "https://ecm.example/", context, 1))
    assert error.value.reason == "wrong_origin"

    async def auth(_page):
        return True

    monkeypatch.setattr(worker, "_page_requires_authentication", auth)
    with pytest.raises(worker.AuthenticationRequired):
        asyncio.run(worker._wait_for_ecm_readiness(page, "https://ecm.example/", context, 10))


def test_criteria_hydration_rejects_empty_shell(monkeypatch):
    async def active(panel):
        return panel

    async def state(_form):
        return {"probe_supported": True, "input_area_present": True, "row_count": 0, "visible_row_count": 0, "missing_properties": ["ECM_Title"]}

    monkeypatch.setattr(worker, "_active_form", active)
    monkeypatch.setattr(worker, "_criteria_hydration_state", state)
    with pytest.raises(worker.CriteriaHydrationError):
        asyncio.run(worker._wait_for_criteria_hydration(object(), 1))


@pytest.mark.parametrize("reason,children,text", [("empty_dom", 0, 0), ("missing_ecm_shell", 1, 2)])
def test_readiness_classifies_shell_failures(monkeypatch, reason, children, text):
    async def state(*_args):
        return {"origin_matches": True, "ready_state": "complete", "body_child_count": children, "body_text_length": text, "shell_counts": {"main": 0}}

    async def no_auth(_page):
        return False

    monkeypatch.setattr(worker, "_ecm_page_state", state)
    monkeypatch.setattr(worker, "_page_requires_authentication", no_auth)
    with pytest.raises(worker.PageNotReadyError) as error:
        asyncio.run(worker._wait_for_ecm_readiness(types.SimpleNamespace(is_closed=lambda: False), "https://ecm.example/", object(), 1))
    assert error.value.reason == reason


def test_safe_reload_once_and_never_on_wrong_origin(monkeypatch):
    class Clock:
        now = 0.0

        def time(self):
            return self.now

    clock = Clock()
    reloads = []
    origin = [True]

    class Page:
        def is_closed(self):
            return False

        async def reload(self, **_kwargs):
            reloads.append(origin[0])

    async def state(*_args):
        return {"origin_matches": origin[0], "ready_state": "complete", "body_child_count": 0, "body_text_length": 0, "shell_counts": {"main": 0}}

    async def no_auth(_page):
        return False

    async def advance(seconds):
        clock.now += seconds

    monkeypatch.setattr(worker.asyncio, "get_running_loop", lambda: clock)
    monkeypatch.setattr(worker.asyncio, "sleep", advance)
    monkeypatch.setattr(worker, "_ecm_page_state", state)
    monkeypatch.setattr(worker, "_page_requires_authentication", no_auth)
    # Create the real event loop before replacing asyncio.get_running_loop.
    loop = asyncio.new_event_loop()
    try:
        with pytest.raises(worker.PageNotReadyError):
            loop.run_until_complete(worker._wait_for_ecm_readiness(Page(), "https://ecm.example/", object(), 4000))
        assert reloads == [True]
        clock.now = 0
        origin[0] = False
        with pytest.raises(worker.PageNotReadyError):
            loop.run_until_complete(worker._wait_for_ecm_readiness(Page(), "https://ecm.example/", object(), 4000))
        assert reloads == [True]
    finally:
        loop.close()


@pytest.mark.parametrize("recent_has_search", [True, False])
def test_saved_search_prefers_recent_then_all(monkeypatch, recent_has_search):
    class Locator:
        def __init__(self, name):
            self.name = name
            self.first = self

        async def wait_for(self, **_kwargs):
            assert opened[0]

    class Page:
        def get_by_role(self, role, name, **_kwargs):
            return Locator(name)

    opened = [False]
    selected = []

    async def visible(locator):
        if locator.name == worker.ECM_SEARCH_NAME:
            return [locator] if opened[0] else []
        return [locator] if locator.name in {"Recent Searches", "All Searches"} else []

    async def expand(locator, _timeout):
        selected.append(locator.name)

    async def saved(_page, group, _search_name):
        return [Locator(group + " search")] if group == "All Searches" or recent_has_search else []

    async def click(locator, _timeout):
        selected.append(locator.name)
        opened[0] = True

    async def hydrate(_panel, _timeout):
        selected.append("hydrated")

    monkeypatch.setattr(worker, "_visible_locators", visible)
    monkeypatch.setattr(worker, "_expand_treeitem", expand)
    monkeypatch.setattr(worker, "_saved_searches_in_group", saved)
    monkeypatch.setattr(worker, "click_control", click)
    monkeypatch.setattr(worker, "_wait_for_criteria_hydration", hydrate)
    asyncio.run(worker._open_saved_search(Page(), 100, search_panel=object()))
    assert selected == (["Recent Searches", "Recent Searches search", "hydrated"] if recent_has_search else ["Recent Searches", "All Searches", "All Searches search", "hydrated"])


def test_other_search_retains_lookup_helper(monkeypatch):
    async def lookup(_page, _timeout, _panel, *, search_name):
        return search_name

    monkeypatch.setattr(worker, "_open_lookup_search", lookup)
    assert asyncio.run(worker._open_saved_search(object(), 10, search_name="Search ECM Document")) == "Search ECM Document"


def test_manual_auth_timeout_and_closed_page(monkeypatch):
    async def requires_auth(_page):
        return True

    monkeypatch.setattr(worker, "_page_requires_authentication", requires_auth)
    with pytest.raises(worker.AuthenticationRequired, match="timed out"):
        asyncio.run(worker._wait_for_manual_authentication(types.SimpleNamespace(is_closed=lambda: False), 0.001))
    with pytest.raises(worker.AuthenticationRequired, match="closed"):
        asyncio.run(worker._wait_for_manual_authentication(types.SimpleNamespace(is_closed=lambda: True), 1))


@pytest.mark.parametrize("headed", [True, False])
def test_authentication_resume_and_headless_fallback(tmp_path, monkeypatch, headed):
    monkeypatch.setattr(worker, "V3_ROOT", tmp_path)
    events = []
    contexts = []
    authenticated = [False]

    class Page:
        url = "https://ecm.example/navigator/"

        async def goto(self, *_args, **_kwargs):
            events.append(("goto", self))

        def is_closed(self):
            return False

    class Context:
        def __init__(self, visible):
            self.visible = visible
            self.pages = [Page()]
            self.closed = False

        async def close(self):
            self.closed = True
            events.append(("close", self.visible))

    class Chromium:
        async def launch_persistent_context(self, **kwargs):
            context = Context(not kwargs["headless"])
            contexts.append(context)
            return context

    class Playwright:
        chromium = Chromium()

        async def __aenter__(self):
            return self

        async def __aexit__(self, *_args):
            pass

    monkeypatch.setitem(sys.modules, "playwright.async_api", types.SimpleNamespace(async_playwright=Playwright))

    async def readiness(page, *_args):
        events.append(("readiness", page))
        if not authenticated[0]:
            raise worker.AuthenticationRequired("login")

    async def manual(page, _timeout):
        events.append(("manual", page))
        authenticated[0] = True

    async def panel(page, *_args):
        events.append(("panel", page))
        return page

    async def search(*_args):
        return 2

    async def export(_page, _panel, temporary, _timeout):
        temporary.write_text("fake CSV")

    def publish(temporary, output, _period):
        output.write_text(temporary.read_text())

    monkeypatch.setattr(worker, "_wait_for_ecm_readiness", readiness)
    monkeypatch.setattr(worker, "_wait_for_manual_authentication", manual)
    monkeypatch.setattr(worker, "_wait_for_search_panel", panel)
    monkeypatch.setattr(worker, "_open_saved_search", panel)
    monkeypatch.setattr(worker, "_search_period", search)
    monkeypatch.setattr(worker, "_export_csv", export)
    monkeypatch.setattr(worker, "_publish_validated_exports", publish)
    result = asyncio.run(worker.run_export(start_date=date(2026, 1, 1), end_date=date(2026, 1, 7), output_path=tmp_path / "out.csv", period_input_path=None, ecm_url="https://ecm.example/", profile_dir=tmp_path / "profile", timeout_seconds=1, root_dir=tmp_path, headed=headed, failure_browser_hold_seconds=0))
    assert result == 2
    assert all(context.closed for context in contexts)
    assert [context.visible for context in contexts] == ([True] if headed else [False, True, False])
    manual_page = next(value for name, value in events if name == "manual")
    if headed:
        assert [value for name, value in events if name == "goto"] == [manual_page]
        assert any(name == "panel" and value is manual_page for name, value in events)


def test_failure_diagnostics_before_close_and_no_handoff(tmp_path, monkeypatch):
    monkeypatch.setattr(worker, "V3_ROOT", tmp_path)
    events = []

    class Page:
        url = "https://ecm.example/?token=secret"

        async def goto(self, *_args, **_kwargs):
            pass

    class Context:
        pages = [Page()]

        async def close(self):
            events.append("close")

    async def launch(**_kwargs):
        return Context()

    class Factory:
        chromium = types.SimpleNamespace(launch_persistent_context=launch)

        async def __aenter__(self):
            return self

        async def __aexit__(self, *_args):
            pass

    async def failure(*_args):
        raise worker.PageNotReadyError("missing_ecm_shell", {})

    async def diagnostics(*_args):
        events.append("diagnostics")

    async def hold(seconds):
        events.append(("hold", seconds))

    monkeypatch.setitem(sys.modules, "playwright.async_api", types.SimpleNamespace(async_playwright=Factory))
    monkeypatch.setattr(worker, "_wait_for_ecm_readiness", failure)
    monkeypatch.setattr(worker, "save_failure_diagnostics", diagnostics)
    monkeypatch.setattr(worker.asyncio, "sleep", hold)
    handoff = tmp_path / "period.csv"
    handoff.write_text("previous")
    with pytest.raises(worker.PageNotReadyError):
        asyncio.run(worker.run_export(start_date=date(2026, 1, 1), end_date=date(2026, 1, 7), output_path=tmp_path / "out.csv", period_input_path=handoff, ecm_url="https://ecm.example/", profile_dir=tmp_path / "profile", timeout_seconds=1, root_dir=tmp_path, diagnostics_dir=tmp_path / "diagnostics", failure_browser_hold_seconds=3))
    assert events == ["diagnostics", ("hold", 3), "close"]
    assert handoff.read_text() == "previous"
    assert not (tmp_path / "out.csv").exists()
