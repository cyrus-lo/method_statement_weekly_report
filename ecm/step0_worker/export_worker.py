"""Export the approved ECM saved search for one reporting period.

The worker deliberately uses a dedicated persistent Edge profile so a user can
complete ECM login/SSO/MFA once without sharing the normal Edge profile. Normal
exports run headlessly from that dedicated profile. It does not use stealth or
anti-detection patches; headed runs may add page-scoped cursor pacing while
using the normal Content Navigator UI.
"""

from __future__ import annotations

import argparse
import asyncio
import csv
import json
import math
import os
import random
import re
import sys
from datetime import date, datetime
from pathlib import Path
from urllib.parse import urlparse


ECM_SEARCH_NAME = "SMM_ECM_Reporting"
CRITERION_PROPERTIES = {
    "OCR Status": "ECM_OCRStatus",
    "Document title": "ECM_Title",
    "Document title:": "ECM_Title",
    "In/Out": "ECM_InOut",
    "Document date": "ECM_DocDate",
}
CRITERION_VALUE_ALIASES = {
    "ECM_OCRStatus": {"0": "Not processed", "1": "Queued", "2": "Processing"},
    "ECM_InOut": {"0": "IN", "1": "BOTH", "2": "OUT"},
}
DEFAULT_ECM_URL = "https://ecm.chunwo.com/navigator/"
CSV_HEADERS = (
    "Document title",
    "Document date",
    "Name",
    "Document revision",
    "Modified By",
    "Major Version",
    "SMM ID",
    "Modified On",
    "Outgoing ref. no.",
    "Path",
    "Link",
)
REQUIRED_ROW_FIELDS = (
    "Document title",
    "Document date",
    "Name",
    "Major Version",
)

CURSOR_TRAVEL_MIN_SECONDS = 0.18
CURSOR_TRAVEL_MAX_SECONDS = 0.70
CURSOR_PAUSE_MIN_SECONDS = 0.04
CURSOR_PAUSE_MAX_SECONDS = 1.50
CURSOR_MIN_STEPS = 8
CURSOR_MAX_STEPS = 24
CURSOR_MAX_INTERMEDIATE_PAUSES = 3


class Step0Error(RuntimeError):
    """Raised when the browser workflow cannot produce a valid export."""


class AuthenticationRequired(Step0Error):
    """Raised when the dedicated Edge profile is not authenticated."""


class PageNotReadyError(Step0Error):
    """Raised when ECM Navigator has not rendered a usable application shell."""

    def __init__(self, reason: str, state: dict[str, object]) -> None:
        self.reason = reason
        self.state = state
        super().__init__(f"ECM Navigator page is not ready ({reason}).")


class CriteriaHydrationError(Step0Error):
    """Raised when a saved-search criteria form remains an empty Dojo shell."""

    def __init__(self, state: dict[str, object]) -> None:
        self.state = state
        super().__init__("ECM saved-search criteria did not finish loading before the timeout.")


class NoSearchResults(Step0Error):
    """Raised when the approved search returns no ECM documents."""


class CsvValidationError(Step0Error):
    """Raised when the downloaded file is not a Step 1-compatible CSV."""


def _log(event: str, **fields: object) -> None:
    """Write a flushed, single-line structured lifecycle message."""
    payload = {"event": event, **fields}
    print(json.dumps(payload, ensure_ascii=True, default=str), flush=True)


def _sanitise_url(value: object) -> str:
    """Keep diagnostic URLs useful without retaining query-string secrets."""
    try:
        parsed = urlparse(str(value))
        return parsed._replace(query="", fragment="").geturl()
    except Exception:
        return "<unavailable>"


def _reset_profile(profile_dir: Path) -> Path | None:
    """Move an existing persistent profile aside before creating a fresh one."""
    if not profile_dir.exists():
        _log("profile_reset_skipped", reason="profile_not_found", profile=profile_dir)
        return None
    timestamp = datetime.now().astimezone().strftime("%Y%m%d-%H%M%S")
    backup_dir = profile_dir.with_name(f"{profile_dir.name}.backup-{timestamp}")
    suffix = 1
    while backup_dir.exists():
        backup_dir = profile_dir.with_name(
            f"{profile_dir.name}.backup-{timestamp}-{suffix}"
        )
        suffix += 1
    try:
        profile_dir.rename(backup_dir)
    except OSError as error:
        raise Step0Error(
            "Could not back up the Edge profile. Close all Edge windows using the "
            f"profile and retry: {profile_dir}"
        ) from error
    _log("profile_reset", previous_profile=profile_dir, backup_profile=backup_dir)
    return backup_dir


class BrowserDiagnostics:
    """Keep a bounded, content-safe record of browser errors for one page."""

    def __init__(self, limit: int = 100) -> None:
        self.limit = limit
        self.console: list[dict[str, object]] = []
        self.page_errors: list[dict[str, object]] = []
        self.request_failures: list[dict[str, object]] = []
        self.response_failures: list[dict[str, object]] = []

    def _append(self, target: list[dict[str, object]], value: dict[str, object]) -> None:
        target.append(value)
        if len(target) > self.limit:
            del target[: len(target) - self.limit]

    def _record(
        self,
        target: list[dict[str, object]],
        event: str,
        value: dict[str, object],
    ) -> None:
        try:
            self._append(target, value)
            try:
                _log(event, **value)
            except Exception:
                # Diagnostics must never interrupt the browser workflow.
                pass
        except Exception:
            # Playwright callbacks must remain non-disruptive even for partial fakes.
            pass

    @staticmethod
    def _console_location(message) -> dict[str, object]:
        try:
            location = getattr(message, "location", {}) or {}
            safe: dict[str, object] = {}
            if isinstance(location, dict):
                if location.get("url"):
                    safe["url"] = _sanitise_url(location["url"])
                for key in ("lineNumber", "columnNumber", "line", "column"):
                    value = location.get(key)
                    if isinstance(value, (int, float)) and not isinstance(value, bool):
                        safe[key] = value
            return safe
        except Exception:
            return {}

    def attach(self, page) -> None:
        if not callable(getattr(page, "on", None)):
            return

        def console(message) -> None:
            try:
                message_type = str(getattr(message, "type", "log"))
                if message_type not in {"warning", "error"}:
                    return
                value = {
                    "type": message_type,
                    "text": str(getattr(message, "text", ""))[:500],
                    "location": self._console_location(message),
                }
                self._record(self.console, "browser_console", value)
            except Exception:
                pass

        def page_error(error) -> None:
            try:
                self._record(
                    self.page_errors,
                    "browser_page_error",
                    {"message": str(error)[:500]},
                )
            except Exception:
                pass

        def request_failed(request) -> None:
            try:
                failure = getattr(request, "failure", None)
                self._record(
                    self.request_failures,
                    "browser_request_failed",
                    {
                        "method": str(getattr(request, "method", "")),
                        "resource_type": str(getattr(request, "resource_type", "")),
                        "url": _sanitise_url(getattr(request, "url", "")),
                        "error": str(failure)[:500],
                    },
                )
            except Exception:
                pass

        def response(response) -> None:
            try:
                status = int(getattr(response, "status", 0) or 0)
                if status < 400:
                    return
                request = getattr(response, "request", None)
                self._record(
                    self.response_failures,
                    "browser_response_failed",
                    {
                        "status": status,
                        "method": str(getattr(request, "method", "")),
                        "url": _sanitise_url(getattr(response, "url", "")),
                    },
                )
            except Exception:
                pass

        page.on("console", console)
        page.on("pageerror", page_error)
        page.on("requestfailed", request_failed)
        page.on("response", response)

    def as_dict(self) -> dict[str, list[dict[str, object]]]:
        return {
            "console": self.console,
            "page_errors": self.page_errors,
            "request_failures": self.request_failures,
            "response_failures": self.response_failures,
        }


async def _save_failure_diagnostics(
    page,
    diagnostics_dir: Path,
    stage: str,
    *,
    browser_diagnostics: BrowserDiagnostics | None = None,
    readiness_state: dict[str, object] | None = None,
) -> None:
    """Save failure evidence without logging browser credentials or page text."""
    diagnostics_dir.mkdir(parents=True, exist_ok=True)
    safe_stage = re.sub(r"[^a-zA-Z0-9_-]+", "_", stage).strip("_") or "unknown"
    screenshot = diagnostics_dir / f"failure-{safe_stage}.png"
    html = diagnostics_dir / f"failure-{safe_stage}.html"
    metadata = diagnostics_dir / f"failure-{safe_stage}.diagnostics.json"
    try:
        await page.screenshot(path=str(screenshot), full_page=True)
        _log("failure_screenshot_saved", path=screenshot)
    except Exception as error:
        _log("failure_screenshot_failed", error=type(error).__name__)
    try:
        html.write_text(await page.content(), encoding="utf-8")
        _log("failure_dom_saved", path=html)
    except Exception as error:
        _log("failure_dom_failed", error=type(error).__name__)
    try:
        title = await page.title()
    except Exception:
        title = "<unavailable>"
    payload = {
        "stage": stage,
        "url": _sanitise_url(getattr(page, "url", "")),
        "title": title,
        "readiness": readiness_state or {},
        "browser_events": browser_diagnostics.as_dict() if browser_diagnostics else {},
    }
    try:
        metadata.write_text(json.dumps(payload, ensure_ascii=True, indent=2, default=str), encoding="utf-8")
        _log("failure_diagnostics_saved", path=metadata)
    except Exception as error:
        _log("failure_diagnostics_failed", error=type(error).__name__)


def parse_iso_date(value: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError as error:
        raise argparse.ArgumentTypeError("Dates must use YYYY-MM-DD format.") from error


def _date_input_value(value: date) -> str:
    return f"{value.month}/{value.day}/{value.year}"


def _parse_ecm_datetime(value: str, field_name: str, row_number: int) -> datetime:
    value = value.strip()
    for date_format in ("%m/%d/%Y, %I:%M %p", "%m/%d/%Y %I:%M %p"):
        try:
            return datetime.strptime(value, date_format)
        except ValueError:
            continue
    raise CsvValidationError(
        f"CSV row {row_number}: {field_name!r} must use M/D/YYYY, H:MM AM/PM; got {value!r}."
    )


def _normalise(value: object) -> str:
    return str(value or "").strip()


def validate_properties_csv(csv_path: Path) -> int:
    """Validate the raw Export All contract and return its data-row count."""
    if not csv_path.is_file():
        raise CsvValidationError(f"Downloaded CSV was not found: {csv_path}")

    with csv_path.open("r", encoding="utf-8-sig", newline="") as source:
        reader = csv.DictReader(source)
        if reader.fieldnames is None:
            raise CsvValidationError("The ECM export does not contain a header row.")
        headers = [_normalise(header) for header in reader.fieldnames]
        missing = [header for header in CSV_HEADERS if header not in headers]
        if missing:
            raise CsvValidationError("ECM export is missing header(s): " + ", ".join(missing))

        row_count = 0
        for row_number, raw_row in enumerate(reader, start=2):
            row = {_normalise(key): _normalise(value) for key, value in raw_row.items() if key}
            if not any(row.values()):
                continue
            row_count += 1
            missing_values = [field for field in REQUIRED_ROW_FIELDS if not row.get(field)]
            if missing_values:
                raise CsvValidationError(
                    f"CSV row {row_number}: required field(s) are blank: "
                    + ", ".join(missing_values)
                )
            _parse_ecm_datetime(row["Document date"], "Document date", row_number)
            if row["Modified On"]:
                _parse_ecm_datetime(row["Modified On"], "Modified On", row_number)
            try:
                int(row["Major Version"])
            except ValueError as error:
                raise CsvValidationError(
                    f"CSV row {row_number}: Major Version must be an integer; "
                    f"got {row['Major Version']!r}."
                ) from error

    if row_count == 0:
        raise NoSearchResults("The ECM export contains no document rows.")
    return row_count


async def _visible_locators(locator) -> list[object]:
    visible = []
    for index in range(await locator.count()):
        candidate = locator.nth(index)
        if await candidate.is_visible():
            visible.append(candidate)
    return visible


class _PageCursor:
    """Pace page-scoped mouse travel without controlling the desktop cursor."""

    def __init__(self, page, *, rng: random.Random | None = None) -> None:
        self.page = page
        self.rng = rng or random.Random()
        self.position: tuple[float, float] | None = None

    def _travel_duration(self, distance: float) -> float:
        """Return a bounded duration that increases with travel distance."""
        distance_factor = min(1.0, max(0.0, distance) / 900.0)
        return CURSOR_TRAVEL_MIN_SECONDS + (
            (CURSOR_TRAVEL_MAX_SECONDS - CURSOR_TRAVEL_MIN_SECONDS)
            * distance_factor
        )

    def _target_from_box(self, box: dict[str, float]) -> tuple[float, float]:
        """Choose a slightly varied point inside a control's bounding box."""
        return (
            float(box["x"])
            + float(box["width"]) * self.rng.uniform(0.35, 0.65),
            float(box["y"])
            + float(box["height"]) * self.rng.uniform(0.35, 0.65),
        )

    def _waypoints(
        self,
        start: tuple[float, float],
        target: tuple[float, float],
    ) -> list[tuple[float, float]]:
        distance = math.hypot(target[0] - start[0], target[1] - start[1])
        step_count = int(
            max(
                CURSOR_MIN_STEPS,
                min(CURSOR_MAX_STEPS, round(distance / 35)),
            )
        )
        return [
            (
                start[0] + (target[0] - start[0]) * (index / step_count),
                start[1] + (target[1] - start[1]) * (index / step_count),
            )
            for index in range(1, step_count + 1)
        ]

    async def move_to_box(self, box: dict[str, float]) -> None:
        """Move through page mouse coordinates before a locator action."""
        target = self._target_from_box(box)
        start = self.position or (0.0, 0.0)
        distance = math.hypot(target[0] - start[0], target[1] - start[1])
        waypoints = self._waypoints(start, target)
        pause_count = self.rng.randint(
            0, min(CURSOR_MAX_INTERMEDIATE_PAUSES, len(waypoints) - 1)
        )
        pause_indices = set(self.rng.sample(range(len(waypoints) - 1), pause_count))
        per_waypoint = self._travel_duration(distance) / len(waypoints)

        for index, (x, y) in enumerate(waypoints):
            await self.page.mouse.move(x, y)
            self.position = (x, y)
            await asyncio.sleep(per_waypoint)
            if index in pause_indices:
                if self.rng.random() < 0.16:
                    pause = self.rng.uniform(0.40, CURSOR_PAUSE_MAX_SECONDS)
                else:
                    pause = self.rng.uniform(CURSOR_PAUSE_MIN_SECONDS, 0.18)
                await asyncio.sleep(min(pause, CURSOR_PAUSE_MAX_SECONDS))


async def _move_cursor_to_locator(locator) -> bool:
    """Move the Playwright page cursor to a locator when its page is available."""
    try:
        page = locator.page
        box = await locator.bounding_box()
        if not box or box.get("width", 0) <= 0 or box.get("height", 0) <= 0:
            return False
        cursor = getattr(page, "_step0_page_cursor", None)
        if cursor is None:
            cursor = _PageCursor(page)
            try:
                setattr(page, "_step0_page_cursor", cursor)
            except Exception:
                pass
        await cursor.move_to_box(box)
        return True
    except Exception:
        # Cursor pacing is deliberately preparatory; the locator action below
        # remains the reliable activation path for Dojo controls.
        return False


async def _click_ui_control(locator, timeout_ms: int, *, force: bool = False) -> None:
    """Click an IBM Dojo control, falling back to its DOM click event."""
    await _move_cursor_to_locator(locator)
    try:
        await locator.click(
            timeout=min(timeout_ms, 10_000), no_wait_after=True, force=force
        )
    except Exception as click_error:
        try:
            await locator.dispatch_event("click")
        except Exception as dispatch_error:
            raise Step0Error(
                "ECM UI control could not be activated: "
                f"{type(click_error).__name__}: {str(click_error)[:240]}"
            ) from dispatch_error


async def _first_visible(locator_candidates: list[object]) -> object | None:
    """Return the first visible locator from a list of selector strategies."""
    for locator in locator_candidates:
        visible = await _visible_locators(locator)
        if visible:
            return visible[0]
    return None


async def _control_metadata(locator) -> dict[str, object]:
    """Return compact metadata for the actual visible control selected by the worker."""
    try:
        return await locator.evaluate(
            """element => ({
                tag: element.tagName,
                id: element.id || '',
                role: element.getAttribute('role') || '',
                aria_label: element.getAttribute('aria-label') || '',
                title: element.getAttribute('title') || '',
                widgetid: element.getAttribute('widgetid') || ''
            })"""
        )
    except Exception:
        return {"metadata": "unavailable"}


async def _search_ui_state(page) -> str:
    """Return safe Search UI diagnostics without recording page contents."""
    try:
        title = await page.title()
    except Exception:
        title = "<unavailable>"
    try:
        url = str(page.url)
    except Exception:
        url = "<unavailable>"
    try:
        panel_count = len(
            await _visible_locators(page.get_by_role("tabpanel", name="Search", exact=True))
        )
    except Exception:
        panel_count = -1
    return f"url={url!r}; title={title!r}; visible_search_panels={panel_count}"


async def _wait_for_search_panel(page, timeout_ms: int):
    """Navigate to and return the Search tab."""
    search_panel = page.get_by_role("tabpanel", name="Search", exact=True)
    deadline = asyncio.get_running_loop().time() + (timeout_ms / 1000)
    last_error: Exception | None = None
    _log("search_navigation_started", timeout_ms=timeout_ms)

    for attempt in range(1, 4):
        if await _visible_locators(search_panel):
            _log("search_panel_ready", source="already_visible")
            return search_panel

        menu_button = await _first_visible(
            [
                page.locator("#icm_widget_SelectorTabContainer_0_tablist_menuBtn"),
                page.locator("[id$='_tablist_menuBtn']"),
                page.get_by_role("button", name=re.compile("menu", re.IGNORECASE)),
            ]
        )
        if menu_button is not None:
            _log("search_menu_button_selected", **await _control_metadata(menu_button))
            await _click_ui_control(menu_button, min(timeout_ms, 10_000))
            await asyncio.sleep(min(0.75, max(0.2, timeout_ms / 1000)))
            _log("search_menu_opened")
        else:
            last_error = Step0Error("ECM feature menu button was not visible.")

        search_item = await _first_visible(
            [
                page.get_by_role("menuitem", name="Search", exact=True),
                page.locator("[role='menuitem']").filter(has_text=re.compile(r"^Search$")),
                page.locator("#dijit_MenuItem_6"),
            ]
        )
        if search_item is None:
            if await _page_requires_authentication(page):
                raise AuthenticationRequired(
                    "The dedicated Edge profile is not authenticated or the ECM Search menu "
                    "did not load. Complete ECM login/MFA before retrying Step 0."
                )
            last_error = Step0Error("ECM Search menu item was not visible after opening the menu.")
        else:
            try:
                _log("search_item_selected", **await _control_metadata(search_item))
                await _click_ui_control(search_item, min(timeout_ms, 10_000))
                remaining_ms = max(
                    500, min(5_000, int((deadline - asyncio.get_running_loop().time()) * 1000))
                )
                await search_panel.first.wait_for(state="visible", timeout=remaining_ms)
                _log("search_panel_ready", source="menu_item")
                return search_panel
            except Exception as error:
                last_error = error

        if asyncio.get_running_loop().time() >= deadline:
            break
        await asyncio.sleep(0.5 * attempt)

    if await _page_requires_authentication(page):
        raise AuthenticationRequired(
            "ECM authentication is required before the Search panel can open."
        ) from last_error
    state = await _search_ui_state(page)
    raise Step0Error(
        "ECM Search panel did not open after selecting Search "
        f"(attempts={attempt}; {state})."
    ) from last_error


async def _page_requires_authentication(page) -> bool:
    """Best-effort detection of a login/SSO page after ECM navigation."""
    try:
        url = str(page.url).casefold()
        title = (await page.title()).casefold()
        body = (await page.locator("body").inner_text()).casefold()
    except Exception:
        return False
    markers = ("login", "sign in", "signin", "password", "single sign-on", "sso")
    return any(marker in f"{url} {title} {body}" for marker in markers)


async def _ecm_page_state(page, ecm_url: str, context) -> dict[str, object]:
    """Return compact Navigator readiness facts without retaining document contents."""
    try:
        document = await page.evaluate(
            """() => {
                const body = document.body;
                return {
                    ready_state: document.readyState,
                    body_text_length: (body?.innerText || '').trim().length,
                    body_html_length: (body?.innerHTML || '').length,
                    body_child_count: body?.childElementCount || 0,
                    iframe_count: document.querySelectorAll('iframe, frame').length,
                };
            }"""
        )
    except Exception:
        document = {
            "ready_state": "<unavailable>",
            "body_text_length": -1,
            "body_html_length": -1,
            "body_child_count": -1,
            "iframe_count": -1,
        }
    try:
        title = await page.title()
    except Exception:
        title = "<unavailable>"
    shell_selectors = (
        "#icm_widget_SelectorTabContainer_0_tablist_menuBtn",
        "[id$='_tablist_menuBtn']",
        "[role='tablist']",
        "[role='main']",
        "[widgetid]",
    )
    shell_counts: dict[str, int] = {}
    for selector in shell_selectors:
        try:
            shell_counts[selector] = await page.locator(selector).count()
        except Exception:
            shell_counts[selector] = -1
    expected_origin = urlparse(ecm_url)
    current_url = _sanitise_url(getattr(page, "url", ""))
    current_origin = urlparse(current_url)
    return {
        "url": current_url,
        "title": title,
        "origin_matches": (
            current_origin.scheme == expected_origin.scheme
            and current_origin.netloc == expected_origin.netloc
        ),
        "page_count": len(getattr(context, "pages", [])),
        "shell_counts": shell_counts,
        **document,
    }


async def _wait_for_ecm_readiness(page, ecm_url: str, context, timeout_ms: int) -> dict[str, object]:
    """Wait for a rendered Navigator shell before interacting with Search controls."""
    loop = asyncio.get_running_loop()
    started = loop.time()
    deadline = started + (timeout_ms / 1000)
    reload_at = started + max(2.0, min(15.0, timeout_ms / 2000))
    reloaded = False
    last_state: dict[str, object] = {}
    _log("page_readiness_started", timeout_ms=timeout_ms)

    while loop.time() < deadline:
        last_state = await _ecm_page_state(page, ecm_url, context)
        shell_count = sum(
            count for count in last_state["shell_counts"].values() if isinstance(count, int) and count > 0
        )
        has_document = (
            int(last_state["body_child_count"]) > 0
            or int(last_state["body_text_length"]) > 0
        )
        if (
            last_state["origin_matches"]
            and last_state["ready_state"] in {"interactive", "complete"}
            and has_document
            and shell_count > 0
        ):
            _log("page_readiness_ready", elapsed_ms=int((loop.time() - started) * 1000), **last_state)
            return last_state

        if await _page_requires_authentication(page):
            _log("page_readiness_authentication_required", **last_state)
            raise AuthenticationRequired("ECM authentication is required before Navigator can load.")

        if not reloaded and loop.time() >= reload_at:
            reloaded = True
            _log("page_readiness_reload", elapsed_ms=int((loop.time() - started) * 1000), **last_state)
            try:
                remaining_ms = max(1_000, int((deadline - loop.time()) * 1000))
                await page.reload(wait_until="domcontentloaded", timeout=remaining_ms)
            except Exception as error:
                _log("page_readiness_reload_failed", error=type(error).__name__)
        await asyncio.sleep(0.5)

    reason = "wrong_origin" if not last_state.get("origin_matches", True) else "empty_dom"
    if last_state.get("body_child_count", 0) or last_state.get("body_text_length", 0):
        reason = "missing_ecm_shell"
    _log("page_readiness_failed", reason=reason, elapsed_ms=int((loop.time() - started) * 1000), **last_state)
    raise PageNotReadyError(reason, last_state)


async def _expand_treeitem(treeitem, timeout_ms: int) -> None:
    """Expand a Dojo tree item when it is not already expanded."""
    expanded = await treeitem.get_attribute("aria-expanded")
    if expanded == "true":
        return
    node = treeitem.locator("xpath=ancestor::div[contains(@class, 'dijitTreeNode')][1]")
    expando = await _visible_locators(node.locator(".dijitTreeExpando"))
    await _click_ui_control(expando[0] if expando else treeitem, timeout_ms)
    deadline = asyncio.get_running_loop().time() + timeout_ms / 1000
    while asyncio.get_running_loop().time() < deadline:
        if await treeitem.get_attribute("aria-expanded") == "true":
            _log("saved_search_group_expanded")
            return
        children = await _visible_locators(node.locator(".dijitTreeNodeContainer [role='treeitem']"))
        if children:
            _log("saved_search_group_expanded", child_count=len(children))
            return
        await asyncio.sleep(0.25)
    raise Step0Error("ECM search tree section did not expand before the timeout.")


async def _saved_searches_in_group(page, group_name: str):
    """Find a saved search using the IBM Dojo tree node, not only ARIA groups."""
    labels = []
    if hasattr(page, "locator"):
        labels = await _visible_locators(
            page.locator(".dijitTreeLabel[role='treeitem']").filter(has_text=group_name)
        )
    if not labels:
        groups = await _visible_locators(page.get_by_role("group", name=group_name, exact=True))
        if len(groups) > 1:
            raise Step0Error(f"ECM search section is ambiguous: {group_name}")
        if not groups:
            return []
        return await _visible_locators(groups[0].get_by_role("treeitem", name=ECM_SEARCH_NAME, exact=True))
    if len(labels) > 1:
        raise Step0Error(f"ECM search section is ambiguous: {group_name}")
    node = labels[0].locator("xpath=ancestor::div[contains(@class, 'dijitTreeNode')][1]")
    candidates = node.locator(".dijitTreeNodeContainer [role='treeitem']").filter(has_text=ECM_SEARCH_NAME)
    return await _visible_locators(candidates)


async def _wait_for_saved_search_tree(page, timeout_ms: int):
    """Wait for Dojo to hydrate the Search tree after its panel becomes visible."""
    _log("saved_search_tree_wait_started", timeout_ms=timeout_ms)
    deadline = asyncio.get_running_loop().time() + timeout_ms / 1000
    while asyncio.get_running_loop().time() < deadline:
        already_open = await _visible_locators(
            page.get_by_role("tab", name=ECM_SEARCH_NAME, exact=True)
        )
        if already_open:
            _log("saved_search_tree_ready", source="already_open", candidate_count=0)
            return
        for group_name in ("Recent Searches", "All Searches"):
            groups = await _visible_locators(
                page.locator(".dijitTreeLabel[role='treeitem']").filter(has_text=group_name)
            )
            if not groups:
                groups = await _visible_locators(
                    page.get_by_role("treeitem", name=group_name, exact=True)
                )
            if groups:
                _log("saved_search_tree_ready", group=group_name, candidate_count=len(groups))
                return
        await asyncio.sleep(0.25)
    raise Step0Error("ECM saved-search tree was not populated before the timeout.")


async def _open_saved_search(page, timeout_ms: int, search_panel=None):
    if search_panel is None:
        search_panel = await _wait_for_search_panel(page, timeout_ms)
    await _wait_for_saved_search_tree(page, timeout_ms)
    recent_searches = page.get_by_role("treeitem", name="Recent Searches", exact=True)
    all_searches = page.get_by_role("treeitem", name="All Searches", exact=True)

    recent_visible = await _visible_locators(recent_searches)
    if len(recent_visible) > 1:
        raise Step0Error("ECM Recent Searches section is ambiguous.")
    if len(recent_visible) == 1:
        await _expand_treeitem(recent_visible[0], timeout_ms)
        saved = await _saved_searches_in_group(page, "Recent Searches")
        if len(saved) == 1:
            _log("saved_search_candidates_found", group="Recent Searches", count=1)
            _log("saved_search_selected", group="Recent Searches", **await _control_metadata(saved[0]))
            await _click_ui_control(saved[0], timeout_ms)
        elif len(saved) > 1:
            raise Step0Error(
                f"Saved ECM search is ambiguous in Recent Searches; found {len(saved)} visible items."
            )
        else:
            all_visible = await _visible_locators(all_searches)
            if len(all_visible) != 1:
                raise Step0Error(f"Saved ECM search was not found: {ECM_SEARCH_NAME}")
            await _expand_treeitem(all_visible[0], timeout_ms)
            saved = await _saved_searches_in_group(page, "All Searches")
            if len(saved) != 1:
                raise Step0Error(f"Saved ECM search was not found: {ECM_SEARCH_NAME}")
            await _click_ui_control(saved[0], timeout_ms)
    else:
        all_visible = await _visible_locators(all_searches)
        if len(all_visible) != 1:
            raise Step0Error(f"Saved ECM search was not found: {ECM_SEARCH_NAME}")
        await _expand_treeitem(all_visible[0], timeout_ms)
        saved = await _saved_searches_in_group(page, "All Searches")
        if len(saved) != 1:
            raise Step0Error(f"Saved ECM search was not found: {ECM_SEARCH_NAME}")
        await _click_ui_control(saved[0], timeout_ms)

    search_tabpanel = page.get_by_role("tabpanel", name=ECM_SEARCH_NAME, exact=True)
    try:
        await search_tabpanel.wait_for(state="visible", timeout=timeout_ms)
        _log("saved_search_tab_ready")
    except Exception as error:
        raise Step0Error(f"Saved ECM search did not open: {ECM_SEARCH_NAME}") from error
    await _wait_for_criteria_hydration(search_tabpanel, timeout_ms)
    return search_tabpanel


async def _select_ecm_page(context, ecm_url: str):
    """Reuse the first existing page on the ECM origin, avoiding duplicates."""
    target = urlparse(ecm_url)
    pages = list(context.pages)
    for page in pages:
        current = urlparse(getattr(page, "url", ""))
        if current.scheme == target.scheme and current.netloc == target.netloc:
            return page
    return pages[0] if pages else await context.new_page()


async def _control_metadata(control) -> dict[str, str]:
    if not hasattr(control, "evaluate"):
        return {"control": "metadata_unavailable"}
    return await control.evaluate(
        """element => {
            const ancestors = [];
            let current = element;
            for (let index = 0; current && index < 5; index += 1, current = current.parentElement) {
                ancestors.push((current.innerText || current.textContent || '').trim());
            }
            const widget = element.closest('[widgetid]');
            const labels = element.labels ? Array.from(element.labels).map(label => label.innerText.trim()) : [];
            const labelledBy = (element.getAttribute('aria-labelledby') || '')
                .split(/\\s+/)
                .filter(Boolean)
                .map(id => document.getElementById(id)?.innerText || document.getElementById(id)?.textContent || '')
                .map(value => value.trim());
            const row = element.closest('tr, .ecmSearchFormFieldArea');
            const rowLabels = row ? Array.from(row.querySelectorAll('label')).map(label => label.innerText.trim()) : [];
            const properties = row ? Array.from(row.querySelectorAll('[name^="ECM_"]')).map(node => node.getAttribute('name') || '') : [];
            return {
                id: element.id || '',
                name: element.getAttribute('name') || '',
                aria: element.getAttribute('aria-label') || '',
                title: element.getAttribute('title') || '',
                alt: element.getAttribute('alt') || '',
                placeholder: element.getAttribute('placeholder') || '',
                role: element.getAttribute('role') || '',
                class_name: element.className || '',
                tag: element.tagName || '',
                widget: element.getAttribute('widgetid') || (widget ? widget.getAttribute('widgetid') : '') || '',
                labels: labels.concat(labelledBy, rowLabels).join(' | '),
                properties: properties.join(' | '),
                context: ancestors.join(' | ')
            };
        }"""
    )


async def _active_criteria_form(search_panel):
    """Return the visible ECM criteria form, avoiding hidden tab duplicates."""
    try:
        candidates = [
            search_panel.locator("[role='search']"),
            search_panel.locator("#ecm_widget_search_SearchForm_0"),
        ]
    except Exception:
        return search_panel
    for candidate in candidates:
        visible = await _visible_locators(candidate)
        if len(visible) == 1:
            return visible[0]
    return search_panel


async def _criteria_hydration_state(criteria_form) -> dict[str, object]:
    """Return compact DOM facts that distinguish a loaded form from its shell."""
    evaluate = getattr(criteria_form, "evaluate", None)
    if not callable(evaluate):
        return {"probe_supported": False}
    return await evaluate(
        """element => {
            const expected = [
                'ECM_DocDate',
                'ECM_OCRStatus',
                'ECM_Title',
                'ECM_InOut',
            ];
            const inputArea = element.querySelector('.ecmSearchFormInputArea');
            const rows = inputArea
                ? Array.from(inputArea.querySelectorAll('tr.ecmSearchFormFieldArea'))
                : [];
            const properties = Array.from(element.querySelectorAll('[name^="ECM_"]'))
                .map(node => node.getAttribute('name') || '')
                .filter(Boolean);
            const uniqueProperties = Array.from(new Set(properties));
            const missingProperties = expected.filter(
                property => !uniqueProperties.includes(property)
            );
            const visibleRows = rows.filter(row => {
                const style = window.getComputedStyle(row);
                const bounds = row.getBoundingClientRect();
                return (
                    style.display !== 'none' &&
                    style.visibility !== 'hidden' &&
                    bounds.width > 0 &&
                    bounds.height > 0
                );
            }).length;
            const searchButton = Array.from(element.querySelectorAll('[role="button"]'))
                .find(button => (button.innerText || '').trim() === 'Search');
            return {
                probe_supported: true,
                input_area_present: Boolean(inputArea),
                row_count: rows.length,
                visible_row_count: visibleRows,
                expected_property_count: expected.length - missingProperties.length,
                missing_properties: missingProperties,
                search_button_disabled: searchButton
                    ? searchButton.getAttribute('aria-disabled') === 'true'
                    : null,
            };
        }"""
    )


async def _wait_for_criteria_hydration(search_panel, timeout_ms: int) -> None:
    """Wait for delayed Dojo criteria rows before resolving saved-search values."""
    loop = asyncio.get_running_loop()
    started = loop.time()
    deadline = started + timeout_ms / 1000
    last_signature: tuple[object, ...] | None = None
    state: dict[str, object] = {}
    _log("criteria_hydration_wait_started", timeout_ms=timeout_ms)

    while loop.time() < deadline:
        criteria_form = await _active_criteria_form(search_panel)
        try:
            state = await _criteria_hydration_state(criteria_form)
        except Exception as error:
            state = {
                "probe_supported": True,
                "probe_error": type(error).__name__,
            }

        if not state.get("probe_supported", True):
            _log(
                "criteria_hydration_ready",
                source="compatibility_fallback",
                elapsed_ms=int((loop.time() - started) * 1000),
            )
            return

        missing = tuple(state.get("missing_properties", ()))
        signature = (
            state.get("input_area_present"),
            state.get("row_count"),
            state.get("visible_row_count"),
            state.get("expected_property_count"),
            missing,
            state.get("search_button_disabled"),
            state.get("probe_error"),
        )
        if signature != last_signature:
            last_signature = signature
            _log(
                "criteria_hydration_state",
                input_area_present=state.get("input_area_present", False),
                row_count=state.get("row_count", 0),
                visible_row_count=state.get("visible_row_count", 0),
                expected_property_count=state.get("expected_property_count", 0),
                missing_properties=list(missing),
                search_button_disabled=state.get("search_button_disabled"),
                probe_error=state.get("probe_error"),
            )

        if (
            state.get("input_area_present")
            and not missing
            and int(state.get("row_count", 0)) >= len(set(CRITERION_PROPERTIES.values()))
            and int(state.get("visible_row_count", 0)) >= len(set(CRITERION_PROPERTIES.values()))
        ):
            _log(
                "criteria_hydration_ready",
                elapsed_ms=int((loop.time() - started) * 1000),
                row_count=state.get("row_count", 0),
                expected_property_count=state.get("expected_property_count", 0),
            )
            return
        await asyncio.sleep(0.25)

    _log(
        "criteria_hydration_failed",
        elapsed_ms=int((loop.time() - started) * 1000),
        last_state=state,
    )
    raise CriteriaHydrationError(state)


async def _labeled_controls(
    search_panel,
    label: str,
    *,
    include_select: bool = True,
    include_operator: bool = False,
):
    selector = "input,textarea,[role='textbox'],[role='combobox']"
    if include_select:
        selector += ",select"
    criteria_form = await _active_criteria_form(search_panel)
    controls = criteria_form.locator(selector)
    matches = []
    seen_keys: set[str] = set()
    wanted = label.casefold().strip()
    property_name = CRITERION_PROPERTIES.get(label)
    property_candidates = []
    if property_name:
        if label != "Document date":
            direct_selectors = [
                f'[name="{property_name}"]:not([type="hidden"])',
                f'[title="{label.rstrip(":")}"]',
                f'[alt="{label.rstrip(":")}"]',
            ]
            direct_locator = criteria_form.locator(", ".join(direct_selectors))
            property_candidates = await _visible_locators(direct_locator)
        if not property_candidates:
            property_rows = criteria_form.locator(
                f'tr:has([name="{property_name}"]), '
                f'.ecmSearchFormFieldArea:has([name="{property_name}"])'
            )
            for row_index in range(await property_rows.count()):
                row = property_rows.nth(row_index)
                row_controls = row.locator(selector)
                property_candidates.extend(await _visible_locators(row_controls))
        if not property_candidates:
            labels = criteria_form.locator("label")
            for label_index in range(await labels.count()):
                label_node = labels.nth(label_index)
                text = (await label_node.inner_text()).strip().rstrip(":").casefold()
                if text != wanted.rstrip(":"):
                    continue
                control_id = await label_node.get_attribute("for")
                if control_id:
                    property_candidates = await _visible_locators(
                        criteria_form.locator(f'#{control_id}')
                    )
                break
    candidate_locators = property_candidates or [
        controls.nth(index) for index in range(await controls.count())
    ]
    for index, control in enumerate(candidate_locators):
        if not await control.is_visible():
            continue
        metadata = await _control_metadata(control)
        searchable = " | ".join(
            metadata[key].casefold()
            for key in ("aria", "title", "alt", "name", "labels", "properties", "context", "widget")
            if metadata.get(key)
        )
        if not property_candidates and wanted not in searchable and (
            not property_name or property_name.casefold() not in searchable
        ):
            continue
        if await control.get_attribute("type") == "hidden":
            continue
        if not include_operator and (
            "operatorsSelect" in metadata.get("class_name", "")
            or "operatorSelect" in metadata.get("class_name", "")
            or "Operator" in metadata.get("id", "")
        ):
            continue
        key = metadata["widget"] or metadata["id"] or f"control-{index}"
        if key not in seen_keys:
            seen_keys.add(key)
            matches.append(control)
    return matches


async def _control_value(control) -> str:
    try:
        value = await control.input_value()
        if value.strip():
            return value.strip()
    except Exception:
        pass
    try:
        value = await control.evaluate(
            """element => {
                const widget = element.closest('[widgetid], .dijitComboBox, .dijitTextBox');
                const row = element.closest('tr, .ecmSearchFormFieldArea');
                const property = row?.querySelector('input[type="hidden"][name^="ECM_"]');
                const candidates = [
                    element.value,
                    element.getAttribute('value'),
                    element.getAttribute('aria-valuetext'),
                    widget?.querySelector('.dijitInputField:not(.dijitValidationContainer)')?.innerText,
                    widget?.innerText,
                    element.parentElement?.innerText,
                    widget?.querySelector('input[type="hidden"]')?.value,
                    property?.value,
                ];
                return candidates.find(value => String(value || '').trim()) || '';
            }"""
        )
        if str(value).strip():
            return str(value).strip()
    except Exception:
        pass
    return (await control.inner_text()).strip()


async def _criterion_value(control, label: str) -> str:
    value = await _control_value(control)
    property_name = CRITERION_PROPERTIES.get(label, "")
    aliases = CRITERION_VALUE_ALIASES.get(property_name, {})
    return aliases.get(value.casefold().strip(), value)


async def _assert_control_value(search_panel, label: str, expected_value: str) -> None:
    controls = await _labeled_controls(search_panel, label)
    property_name = CRITERION_PROPERTIES.get(label, "")
    _log(
        "criteria_controls_discovered",
        criterion=label,
        property=property_name,
        candidate_count=len(controls),
    )
    if not controls:
        raise Step0Error(f"Saved search field was not found: {label}")
    values = [(control, await _criterion_value(control, label)) for control in controls]
    _log(
        "criteria_values_observed",
        criterion=label,
        property=property_name,
        values=[value[:120] for _, value in values],
    )
    matching = [
        control
        for control, value in values
        if value.casefold().strip() == expected_value.casefold().strip()
        or expected_value.casefold().strip() in value.casefold().strip()
    ]
    if len(matching) != 1:
        actual = ", ".join(repr(value) for _, value in values)
        if len(matching) > 1:
            raise Step0Error(f"Saved search field is ambiguous: {label}")
        raise Step0Error(
            f"Saved search criterion {label!r} is {actual or '<blank>'}; "
            f"expected {expected_value!r}."
        )


async def _assert_saved_search_criteria(search_panel) -> None:
    expected = (
        ("OCR Status", "Processing"),
        ("Document title:", "Method Statement"),
        ("In/Out", "OUT"),
    )
    for label, expected_value in expected:
        await _assert_control_value(search_panel, label, expected_value)


async def _search_period(page, search_panel, start_date: date, end_date: date, timeout_ms: int) -> int:
    await _assert_saved_search_criteria(search_panel)

    operator_controls = await _labeled_controls(
        search_panel,
        "Document date",
        include_select=True,
        include_operator=True,
    )
    operator_matches = []
    for control in operator_controls:
        metadata = await _control_metadata(control)
        if "Operator" not in str(metadata.get("id", "")):
            continue
        if (await _control_value(control)).casefold() == "between":
            operator_matches.append(control)
    if len(operator_matches) != 1:
        raise Step0Error("Document date operator 'Between' was not found in the saved search.")

    date_inputs = search_panel.locator('input[placeholder*="M/d/yyyy"]')
    visible_date_inputs = await _visible_locators(date_inputs)
    if len(visible_date_inputs) != 2:
        date_inputs = search_panel.locator("input")
        visible_date_inputs = []
        for index in range(await date_inputs.count()):
            candidate = date_inputs.nth(index)
            if not await candidate.is_visible():
                continue
            metadata = await _control_metadata(candidate)
            searchable = " ".join(
                metadata[key].casefold() for key in ("aria", "name", "placeholder", "context")
            )
            if "document date" in searchable or "m/d/yyyy" in searchable:
                visible_date_inputs.append(candidate)
    if len(visible_date_inputs) != 2:
        raise Step0Error("Expected exactly two Document date inputs in the saved search.")
    await _move_cursor_to_locator(visible_date_inputs[0])
    await visible_date_inputs[0].fill(_date_input_value(start_date))
    await _move_cursor_to_locator(visible_date_inputs[1])
    await visible_date_inputs[1].fill(_date_input_value(end_date))
    await visible_date_inputs[1].dispatch_event("change")
    await visible_date_inputs[1].press("Tab")

    search_button = search_panel.get_by_role("button", name="Search", exact=True)
    visible_search_buttons = await _visible_locators(search_button)
    if len(visible_search_buttons) != 1:
        raise Step0Error("ECM Search button was not found.")
    await _click_ui_control(visible_search_buttons[0], timeout_ms)

    results_toolbar = search_panel.get_by_role("toolbar", name="SearchResultsToolbar")
    try:
        await results_toolbar.first.wait_for(state="visible", timeout=timeout_ms)
    except Exception as error:
        raise Step0Error("ECM Search Results did not finish loading before the timeout.") from error

    body_text = await page.locator("body").inner_text()
    count_match = re.search(r"The search returned\s+([\d,]+)\s+items?", body_text, re.IGNORECASE)
    if count_match:
        result_count = int(count_match.group(1).replace(",", ""))
    else:
        grid = search_panel.get_by_role("grid", name="grid")
        row_count = await grid.get_by_role("row").count()
        result_count = max(row_count - 1, 0)

    if result_count == 0:
        raise NoSearchResults(
            f"The ECM search returned no documents for {start_date.isoformat()} to "
            f"{end_date.isoformat()}."
        )
    return result_count


async def _export_csv(page, search_panel, temporary_path: Path, timeout_ms: int) -> None:
    export_button = search_panel.get_by_role("button", name="Export All", exact=True)
    visible_export_buttons = await _visible_locators(export_button)
    if len(visible_export_buttons) != 1:
        export_button = page.get_by_role("button", name="Export All", exact=True)
        visible_export_buttons = await _visible_locators(export_button)
    if len(visible_export_buttons) != 1:
        raise Step0Error("ECM Export All button was not found.")

    await _click_ui_control(visible_export_buttons[0], timeout_ms)
    export_dialog = page.get_by_role("dialog", name="Export All", exact=True)
    try:
        await export_dialog.first.wait_for(
            state="visible", timeout=min(timeout_ms, 10_000)
        )
    except Exception as error:
        raise Step0Error("ECM Export All dialog did not open.") from error

    visible_dialogs = await _visible_locators(export_dialog)
    if len(visible_dialogs) != 1:
        raise Step0Error("ECM Export All dialog is ambiguous.")
    export_dialog = visible_dialogs[0]

    all_items = export_dialog.get_by_role("radio", name="Export all items", exact=True)
    visible_all_items = await _visible_locators(all_items)
    if len(visible_all_items) != 1:
        raise Step0Error("ECM Export All dialog did not offer 'Export all items'.")
    if not await visible_all_items[0].is_checked():
        await _click_ui_control(visible_all_items[0], timeout_ms)

    raw_values = export_dialog.get_by_role(
        "checkbox", name="Export raw values", exact=True
    )
    visible_raw_values = await _visible_locators(raw_values)
    if len(visible_raw_values) != 1:
        raise Step0Error("ECM Export All dialog did not offer the raw-values option.")
    if await visible_raw_values[0].is_checked():
        raise Step0Error("ECM Export All raw values must remain unchecked.")

    export_action = export_dialog.get_by_role("button", name="Export", exact=True)
    visible_export_actions = await _visible_locators(export_action)
    if len(visible_export_actions) != 1:
        raise Step0Error("ECM Export All confirmation button was not found.")

    temporary_path.parent.mkdir(parents=True, exist_ok=True)
    async with page.expect_download(timeout=timeout_ms) as download_info:
        await _click_ui_control(visible_export_actions[0], timeout_ms)
    download = await download_info.value
    failure = await download.failure()
    if failure:
        raise Step0Error(f"ECM export download failed: {failure}")
    await download.save_as(str(temporary_path))


async def _wait_for_authentication(
    playwright,
    ecm_url: str,
    profile_dir: Path,
    timeout_seconds: int,
) -> None:
    """Show the dedicated profile while the user completes ECM authentication."""
    _log("waiting_for_authentication")
    context = await playwright.chromium.launch_persistent_context(
        user_data_dir=str(profile_dir),
        channel="msedge",
        headless=False,
        accept_downloads=True,
    )
    deadline = asyncio.get_running_loop().time() + timeout_seconds
    try:
        page = await _select_ecm_page(context, ecm_url)
        await page.goto(ecm_url, wait_until="domcontentloaded", timeout=timeout_seconds * 1000)
        while asyncio.get_running_loop().time() < deadline:
            try:
                await _open_saved_search(page, 2_000)
            except AuthenticationRequired:
                await asyncio.sleep(1)
                continue
            except Step0Error:
                # The login may be complete even when the Search UI is still loading.
                await asyncio.sleep(1)
                continue
            _log("authentication_detected")
            return
        raise AuthenticationRequired(
            f"ECM authentication was not completed within {timeout_seconds} seconds."
        )
    finally:
        await context.close()


async def run_export(
    *,
    start_date: date,
    end_date: date,
    output_path: Path,
    period_input_path: Path | None = None,
    ecm_url: str,
    profile_dir: Path,
    timeout_seconds: int,
    headed: bool = False,
    login_if_needed: bool = False,
    auth_timeout_seconds: int = 120,
    failure_browser_hold_seconds: int = 0,
    reset_profile: bool = False,
) -> int:
    """Run the saved ECM search and atomically publish a validated CSV."""
    if output_path.exists():
        raise Step0Error(f"Refusing to overwrite an existing CSV: {output_path}")
    if reset_profile:
        _reset_profile(profile_dir)
    profile_dir.mkdir(parents=True, exist_ok=True)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = output_path.with_name(f".{output_path.name}.download")
    if temporary_path.exists():
        temporary_path.unlink()
    diagnostics_dir = Path(
        os.getenv("STEP0_DIAGNOSTICS_DIR", str(output_path.parent / ".step0-diagnostics"))
    )

    try:
        from playwright.async_api import async_playwright
    except ImportError as error:
        raise Step0Error(
            "Playwright is not installed in the Step 0 worker environment. "
            "Install ecm/step0_worker/requirements.txt first."
        ) from error

    context = None
    stage = "starting ECM browser"
    try:
        async with async_playwright() as playwright:
            resume_current_page = False
            while True:
                if context is None:
                    context = await playwright.chromium.launch_persistent_context(
                        user_data_dir=str(profile_dir),
                        channel="msedge",
                        headless=not headed,
                        accept_downloads=True,
                    )
                try:
                    page = await _select_ecm_page(context, ecm_url)
                    browser_diagnostics = BrowserDiagnostics()
                    browser_diagnostics.attach(page)
                    readiness_state: dict[str, object] | None = None
                    stage = "opening ECM Navigator"
                    _log("stage_started", stage=stage)
                    if resume_current_page:
                        _log("resuming_on_authenticated_page")
                        resume_current_page = False
                    else:
                        await page.goto(
                            ecm_url,
                            wait_until="domcontentloaded",
                            timeout=timeout_seconds * 1000,
                        )
                    if login_if_needed and not headed and await _page_requires_authentication(page):
                        raise AuthenticationRequired(
                            "ECM authentication is required before the Search page can open."
                        )
                    stage = "waiting for ECM Navigator readiness"
                    _log("stage_started", stage=stage)
                    readiness_state = await _wait_for_ecm_readiness(
                        page, ecm_url, context, timeout_seconds * 1000
                    )
                    stage = "opening the Search page"
                    _log("stage_started", stage=stage)
                    if hasattr(page, "get_by_role"):
                        search_panel = await _wait_for_search_panel(
                            page, timeout_seconds * 1000
                        )
                    stage = "opening the saved search"
                    _log("stage_started", stage=stage)
                    search_panel = await _open_saved_search(page, timeout_seconds * 1000)
                    stage = "running the saved search"
                    _log("stage_started", stage=stage)
                    result_count = await _search_period(
                        page, search_panel, start_date, end_date, timeout_seconds * 1000
                    )
                    stage = "downloading Export All"
                    _log("stage_started", stage=stage)
                    await _export_csv(page, search_panel, temporary_path, timeout_seconds * 1000)
                    break
                except AuthenticationRequired:
                    if not login_if_needed:
                        raise
                    if headed:
                        _log("waiting_for_authentication")
                        deadline = asyncio.get_running_loop().time() + auth_timeout_seconds
                        while asyncio.get_running_loop().time() < deadline:
                            if not await _page_requires_authentication(page):
                                _log("authentication_detected")
                                break
                            await asyncio.sleep(1)
                        else:
                            raise AuthenticationRequired(
                                "ECM authentication was not completed within "
                                f"{auth_timeout_seconds} seconds."
                            )
                        _log("resuming_export_after_authentication")
                        resume_current_page = True
                        continue
                    _log("authentication_required")
                    await context.close()
                    context = None
                    await _wait_for_authentication(
                        playwright,
                        ecm_url,
                        profile_dir,
                        auth_timeout_seconds,
                    )
                    _log("resuming_headless_export")
                except Exception as error:
                    _log(
                        "stage_failed",
                        stage=stage,
                        error=type(error).__name__,
                        message=str(error)[:500],
                    )
                    try:
                        failed_readiness_state = (
                            error.state if isinstance(error, PageNotReadyError) else readiness_state
                        )
                        _log(
                            "page_state",
                            stage=stage,
                            url=_sanitise_url(page.url),
                            title=await page.title(),
                        )
                        await _save_failure_diagnostics(
                            page,
                            diagnostics_dir,
                            stage,
                            browser_diagnostics=browser_diagnostics,
                            readiness_state=failed_readiness_state,
                        )
                    except Exception:
                        pass
                    if headed and failure_browser_hold_seconds > 0:
                        try:
                            _log(
                                "browser_retained_for_inspection",
                                seconds=failure_browser_hold_seconds,
                            )
                            await asyncio.sleep(failure_browser_hold_seconds)
                        except asyncio.CancelledError:
                            raise
                    raise error
                finally:
                    if context is not None and not resume_current_page:
                        await context.close()
                        context = None

            _log("stage_started", stage="validating CSV")
            validate_properties_csv(temporary_path)
            os.replace(temporary_path, output_path)
            _log("stage_started", stage="publishing CSV")
            if period_input_path is not None:
                period_temporary_path = period_input_path.with_name(
                    f".{period_input_path.name}.download"
                )
                try:
                    import shutil

                    shutil.copy2(output_path, period_temporary_path)
                    os.replace(period_temporary_path, period_input_path)
                finally:
                    if period_temporary_path.exists():
                        period_temporary_path.unlink()
            _log("workflow_complete", result_count=result_count)
            return result_count
    finally:
        if temporary_path.exists():
            temporary_path.unlink()


async def login(profile_dir: Path, ecm_url: str, timeout_seconds: int) -> None:
    """Open a dedicated visible Edge profile for manual login/SSO/MFA."""
    profile_dir.mkdir(parents=True, exist_ok=True)
    try:
        from playwright.async_api import async_playwright
    except ImportError as error:
        raise Step0Error(
            "Playwright is not installed in the Step 0 worker environment. "
            "Install ecm/step0_worker/requirements.txt first."
        ) from error

    async with async_playwright() as playwright:
        context = await playwright.chromium.launch_persistent_context(
            user_data_dir=str(profile_dir),
            channel="msedge",
            headless=False,
            accept_downloads=True,
        )
        try:
            page = context.pages[0] if context.pages else await context.new_page()
            await page.goto(ecm_url, wait_until="domcontentloaded", timeout=timeout_seconds * 1000)
            print(
                "Complete ECM login/SSO/MFA in the dedicated Edge window. "
                "When the ECM page is ready, return here and press Enter."
            )
            await asyncio.to_thread(input)
        finally:
            await context.close()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Export the ECM SMM_ECM_Reporting saved search.")
    parser.add_argument("--start-date", type=parse_iso_date, required=False)
    parser.add_argument("--end-date", type=parse_iso_date, required=False)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--period-input", type=Path)
    parser.add_argument("--ecm-url", default=os.getenv("ECM_BASE_URL", DEFAULT_ECM_URL))
    parser.add_argument("--profile-dir", type=Path, required=True)
    parser.add_argument("--timeout-seconds", type=int, default=60)
    parser.add_argument(
        "--login-if-needed",
        action="store_true",
        help="Open a visible dedicated Edge window only when ECM authentication is required.",
    )
    parser.add_argument(
        "--auth-timeout-seconds",
        type=int,
        default=120,
        help="Maximum time to wait for interactive ECM authentication.",
    )
    parser.add_argument("--login", action="store_true", help="Open visible Edge for manual login/SSO/MFA.")
    parser.add_argument("--headed", action="store_true", help="Run export with the dedicated Edge visible.")
    parser.add_argument(
        "--failure-browser-hold-seconds",
        type=int,
        default=0,
        help="Keep a headed browser open after failure for inspection; zero disables retention.",
    )
    parser.add_argument(
        "--reset-profile",
        action="store_true",
        help="Back up the persistent Edge profile and start with a fresh profile for this run.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.timeout_seconds <= 0:
        raise Step0Error("--timeout-seconds must be greater than zero.")
    if args.auth_timeout_seconds <= 0:
        raise Step0Error("--auth-timeout-seconds must be greater than zero.")
    if args.failure_browser_hold_seconds < 0:
        raise Step0Error("--failure-browser-hold-seconds must not be negative.")
    if args.login:
        asyncio.run(login(args.profile_dir, args.ecm_url, args.timeout_seconds))
        return 0
    if args.start_date is None or args.end_date is None or args.output is None:
        raise Step0Error("--start-date, --end-date, and --output are required for export mode.")
    if args.end_date < args.start_date:
        raise Step0Error("End date cannot be earlier than start date.")
    result_count = asyncio.run(
        run_export(
            start_date=args.start_date,
            end_date=args.end_date,
            output_path=args.output,
            period_input_path=args.period_input,
            ecm_url=args.ecm_url,
            profile_dir=args.profile_dir,
            timeout_seconds=args.timeout_seconds,
            headed=args.headed,
            login_if_needed=args.login_if_needed,
            auth_timeout_seconds=args.auth_timeout_seconds,
            failure_browser_hold_seconds=args.failure_browser_hold_seconds,
            reset_profile=args.reset_profile,
        )
    )
    print(f"Search: {ECM_SEARCH_NAME}")
    print(f"Search results: {result_count}")
    print(f"Output CSV: {args.output.resolve()}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (CsvValidationError, Step0Error, OSError, ValueError) as error:
        print(f"Error: {error}", file=sys.stderr)
        raise SystemExit(1)
