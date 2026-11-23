"""V3-owned ECM saved-search CSV export worker.

Only V3 code may write outputs, diagnostics, or browser-profile state. This worker
uses the standard Content Navigator UI and a persistent, dedicated Edge profile.
"""

from __future__ import annotations

import argparse
import asyncio
import csv
import json
import os
import re
import shutil
import sys
from datetime import date, datetime
from pathlib import Path
from urllib.parse import urlparse

from .browser_support import BrowserDiagnostics, click_control, reset_profile, save_failure_diagnostics
from .properties_csv import CSV_HEADERS, CsvValidationError, validate_properties_csv

ECM_SEARCH_NAME = "SMM_ECM_Reporting"
CRITERION_PROPERTIES = {
    "OCR Status": "ECM_OCRStatus",
    "Document title": "ECM_Title",
    "Document title:": "ECM_Title",
    "In/Out": "ECM_InOut",
    "Document date": "ECM_DocDate",
    "Outgoing ref. no.": "ECM_OutRefNo",
}
CRITERION_VALUE_ALIASES = {"ECM_OCRStatus": {"0": "Not processed", "1": "Queued", "2": "Processing"}, "ECM_InOut": {"0": "IN", "1": "BOTH", "2": "OUT"}}
DEFAULT_ECM_URL = "https://ecm.chunwo.com/navigator/"
V3_ROOT = Path(__file__).resolve().parents[1]


class Step0Error(RuntimeError):
    """The ECM export could not safely complete."""


class NoSearchResults(Step0Error):
    """The saved search returned no exportable records."""


class AuthenticationRequired(Step0Error):
    """Authentication is required or did not complete in time."""


class PageNotReadyError(Step0Error):
    """Navigator did not render a usable shell."""

    def __init__(self, reason: str, state: dict[str, object]) -> None:
        self.reason, self.state = reason, state
        super().__init__(f"ECM Navigator page is not ready ({reason}).")


class CriteriaHydrationError(Step0Error):
    """The saved search did not populate its criteria."""

    def __init__(self, state: dict[str, object]) -> None:
        self.state = state
        super().__init__("ECM saved-search criteria did not finish loading before the timeout.")


def _log(event: str, **fields: object) -> None:
    print(json.dumps({"event": event, **fields}, ensure_ascii=True, default=str), flush=True)


def _inside(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root.resolve())
        return True
    except ValueError:
        return False


def _date_input_value(value: date) -> str:
    return f"{value.month}/{value.day}/{value.year}"


def _sanitise_url(value: object) -> str:
    try:
        return urlparse(str(value))._replace(query="", fragment="").geturl()
    except Exception:
        return "<unavailable>"


def _publish_validated_exports(temporary: Path, output_path: Path, period_input_path: Path | None) -> None:
    """Validate and publish both Step 0 CSVs without leaving a partial new export."""
    validate_properties_csv(temporary)
    staged_output = output_path.with_name(f".{output_path.name}.publish")
    staged_period: Path | None = None
    published_output = False
    try:
        shutil.copy2(temporary, staged_output)
        validate_properties_csv(staged_output)
        if period_input_path:
            period_input_path.parent.mkdir(parents=True, exist_ok=True)
            staged_period = period_input_path.with_name(f".{period_input_path.name}.publish")
            shutil.copy2(staged_output, staged_period)
            validate_properties_csv(staged_period)
        os.replace(staged_output, output_path)
        published_output = True
        if period_input_path and staged_period:
            os.replace(staged_period, period_input_path)
    except Exception:
        if published_output:
            output_path.unlink(missing_ok=True)
        raise
    finally:
        staged_output.unlink(missing_ok=True)
        if staged_period:
            staged_period.unlink(missing_ok=True)


async def _visible_locators(locator) -> list:
    visible = []
    for index in range(await locator.count()):
        candidate = locator.nth(index)
        if await candidate.is_visible():
            visible.append(candidate)
    return visible


async def _control_metadata(control) -> dict[str, str]:
    return await control.evaluate(
        """element => {
            const ancestors = [];
            let current = element;
            for (let index = 0; current && index < 5; index += 1, current = current.parentElement) {
                ancestors.push((current.innerText || current.textContent || '').trim());
            }
            const widget = element.closest('[widgetid]');
            const row = element.closest('tr, .ecmSearchFormFieldArea');
            const labels = element.labels ? Array.from(element.labels).map(label => label.innerText.trim()) : [];
            return {
                id: element.id || '', name: element.getAttribute('name') || '',
                aria: element.getAttribute('aria-label') || '', title: element.getAttribute('title') || '',
                alt: element.getAttribute('alt') || '', placeholder: element.getAttribute('placeholder') || '',
                class_name: element.className || '', widget: element.getAttribute('widgetid') || (widget ? widget.getAttribute('widgetid') : '') || '',
                labels: labels.concat((element.getAttribute('aria-labelledby') || '').split(/\\s+/).filter(Boolean).map(id => document.getElementById(id)?.innerText || ''), row ? Array.from(row.querySelectorAll('label')).map(node => node.innerText.trim()) : []).join(' | '), properties: row ? Array.from(row.querySelectorAll('[name^="ECM_"]')).map(node => node.getAttribute('name') || '').join(' | ') : '',
                context: ancestors.join(' | ')
            };
        }"""
    )


async def _labeled_controls(form, label: str, *, include_select: bool = True, include_operator: bool = False) -> list:
    selector = "input,textarea,[role='textbox'],[role='combobox']" + (",select" if include_select else "")
    controls = form.locator(selector)
    property_name = CRITERION_PROPERTIES.get(label)
    candidates = []
    if property_name:
        if label != "Document date" and not include_operator:
            direct = form.locator(
                f'[name="{property_name}"]:not([type="hidden"]),[title="{label.rstrip(":")}"],[alt="{label.rstrip(":")}"]'
            )
            candidates = await _visible_locators(direct)
        if not candidates:
            rows = form.locator(f'tr:has([name="{property_name}"]), .ecmSearchFormFieldArea:has([name="{property_name}"])')
            for index in range(await rows.count()):
                candidates.extend(await _visible_locators(rows.nth(index).locator(selector)))
        if not candidates:
            labels = form.locator("label")
            for index in range(await labels.count()):
                node = labels.nth(index)
                if (await node.inner_text()).strip().rstrip(":").casefold() == label.rstrip(":").casefold():
                    control_id = await node.get_attribute("for")
                    if control_id:
                        candidates = await _visible_locators(form.locator(f'#{control_id}'))
                    break
    property_candidates_found = bool(candidates)
    if not candidates:
        candidates = [controls.nth(index) for index in range(await controls.count())]
    wanted = label.casefold().rstrip(":")
    result = []
    seen: set[str] = set()
    for index, control in enumerate(candidates):
        if not await control.is_visible():
            continue
        control_type = (await control.get_attribute("type") or "").casefold()
        if control_type in {"hidden", "button", "submit", "reset"} or await control.get_attribute("aria-hidden") == "true":
            continue
        metadata = await _control_metadata(control)
        class_name = metadata.get("class_name", "")
        if any(marker in class_name for marker in ("dijitOffScreen", "dijitValidationInner", "dijitValidationIcon", "dijitArrowButtonInner")):
            continue
        searchable = " | ".join(metadata.get(key, "") for key in ("aria", "title", "alt", "name", "labels", "properties", "context", "widget")).casefold()
        if not property_candidates_found and wanted not in searchable and not (property_name and property_name.casefold() in searchable):
            # A direct property selector may not expose its property name in parent metadata.
            continue
        if not include_operator and ("operatorsSelect" in metadata.get("class_name", "") or "operatorSelect" in metadata.get("class_name", "") or "Operator" in metadata.get("id", "")):
            continue
        key = metadata.get("widget") or metadata.get("id") or f"control-{index}"
        if key not in seen:
            seen.add(key)
            result.append(control)
    return result


async def _step0_control_value(control) -> str:
    try:
        value = await control.input_value()
        if str(value).strip():
            return str(value).strip()
    except Exception:
        pass
    try:
        value = await control.evaluate(
            """element => {
                const widget = element.closest('[widgetid], .dijitComboBox, .dijitTextBox');
                const row = element.closest('tr, .ecmSearchFormFieldArea');
                const property = row?.querySelector('input[type="hidden"][name^="ECM_"]');
                return [element.value, element.getAttribute('value'), element.getAttribute('aria-valuetext'),
                    widget?.querySelector('.dijitInputField:not(.dijitValidationContainer)')?.innerText,
                    widget?.innerText, element.parentElement?.innerText, widget?.querySelector('input[type="hidden"]')?.value, property?.value]
                    .find(value => String(value || '').trim()) || '';
            }"""
        )
        if str(value).strip():
            return str(value).strip()
    except Exception:
        pass
    try:
        return str(await control.inner_text()).strip()
    except Exception:
        return ""


async def _control_value(control) -> str:
    # Preserve the shared lookup worker's existing value precedence.
    try:
        value = await control.input_value()
        if str(value).strip():
            return str(value).strip()
    except Exception:
        pass
    try:
        value = await control.inner_text()
        if str(value).strip():
            return str(value).strip()
    except Exception:
        pass
    return await _step0_control_value(control)


async def _criterion_value(control, label: str) -> str:
    value = await _step0_control_value(control)
    aliases = CRITERION_VALUE_ALIASES.get(CRITERION_PROPERTIES.get(label, ""), {})
    return aliases.get(value.casefold(), value)


async def _assert_value(form, label: str, expected: str) -> None:
    controls = await _labeled_controls(form, label)
    values = [await _criterion_value(control, label) for control in controls]
    matches = [value for value in values if expected.casefold() == value.strip().casefold()]
    _log("criteria_values", criterion=label, values=values)
    if len(matches) != 1:
        raise Step0Error(f"Saved search criterion {label!r} must be exactly {expected!r}; observed {values!r}.")


async def _active_form(search_panel):
    for selector in ("[role='search']", "#ecm_widget_search_SearchForm_0"):
        visible = await _visible_locators(search_panel.locator(selector))
        if len(visible) == 1:
            return visible[0]
    return search_panel


async def _ensure_lookup_criteria_expanded(panel, timeout_ms: int) -> None:
    if not callable(getattr(panel, "locator", None)):
        return
    try:
        collapsed = await panel.locator(".searchToggleArea.dijitHidden").count() > 0
        if not collapsed:
            header = panel.locator('[data-dojo-attach-point="searchCriteriaHeaderTitleNode"][aria-pressed="false"]')
            collapsed = await header.count() > 0
    except Exception:
        return
    if not collapsed:
        return
    target = await _first_visible((
        panel.locator('[data-dojo-attach-point="titleToggleNode"]'),
        panel.locator('[data-dojo-attach-point="imageToggleNode"]'),
        panel.locator('[data-dojo-attach-point="searchCriteriaHeaderTitleNode"]'),
    ))
    if target is None:
        return
    _log("lookup_criteria_expanding")
    await click_control(target, min(timeout_ms, 10_000))
    deadline = asyncio.get_running_loop().time() + timeout_ms / 1000
    while asyncio.get_running_loop().time() < deadline:
        try:
            if await panel.locator(".searchToggleArea.dijitHidden").count() == 0:
                return
        except Exception:
            return
        await asyncio.sleep(0.1)
    raise Step0Error("ECM Search Criteria section remained collapsed after clicking its header.")


async def _wait_for_lookup_criteria_hydration(panel, timeout_ms: int) -> None:
    deadline = asyncio.get_running_loop().time() + timeout_ms / 1000
    state: dict[str, object] = {}
    while asyncio.get_running_loop().time() < deadline:
        form = await _active_form(panel)
        if not callable(getattr(form, "evaluate", None)):
            return
        try:
            state = await form.evaluate("""element => {
                const area = element.querySelector('.ecmSearchFormInputArea');
                const rows = area ? Array.from(area.querySelectorAll('tr.ecmSearchFormFieldArea')) : [];
                const properties = Array.from(element.querySelectorAll('[name^="ECM_"]')).map(node => node.getAttribute('name'));
                const required = ['ECM_InOut', 'ECM_OutRefNo'];
                const visibleRows = rows.filter(row => {
                    const style = window.getComputedStyle(row), box = row.getBoundingClientRect();
                    return style.display !== 'none' && style.visibility !== 'hidden' && box.width > 0 && box.height > 0;
                }).length;
                return {input_area_present: Boolean(area), row_count: rows.length, visible_row_count: visibleRows,
                    missing_properties: required.filter(name => !properties.includes(name))};
            }""")
        except Exception:
            state = {}
        if (state.get("input_area_present") and not state.get("missing_properties")
                and state.get("row_count", 0) >= 2 and state.get("visible_row_count", 0) >= 2):
            return
        await asyncio.sleep(0.25)
    _log("lookup_criteria_hydration_failed", state=state)
    raise Step0Error("ECM lookup criteria did not finish loading before the timeout.")


async def _open_lookup_search(page, timeout_ms: int, search_panel=None, *, search_name: str):
    search_panel_locator = page.get_by_role("tabpanel", name=search_name, exact=True)
    visible_panel = await _visible_locators(search_panel_locator)
    if len(visible_panel) > 1:
        raise Step0Error("Saved ECM search panel is ambiguous.")
    if not visible_panel:
        _log("saved_search_navigation", search=search_name)
        if search_panel is None:
            search_panel = await _wait_for_search_panel(page, timeout_ms)
        deadline = asyncio.get_running_loop().time() + timeout_ms / 1000
        groups = {}
        while asyncio.get_running_loop().time() < deadline:
            for name in ("Recent Searches", "All Searches"):
                groups[name] = await _visible_locators(page.get_by_role("treeitem", name=name, exact=True))
            if any(groups.values()):
                break
            if await _visible_locators(search_panel_locator):
                break
            await asyncio.sleep(0.25)
        if not await _visible_locators(search_panel_locator):
            for name in ("Recent Searches", "All Searches"):
                matches = groups.get(name, [])
                if len(matches) > 1:
                    raise Step0Error(f"ECM search section is ambiguous: {name}")
                if not matches:
                    continue
                await _expand_treeitem(matches[0], timeout_ms)
                saved = await _saved_searches_in_group(page, name, search_name)
                if len(saved) > 1:
                    raise Step0Error(f"Saved ECM search is ambiguous in {name}.")
                if saved:
                    await click_control(saved[0], timeout_ms)
                    break
            else:
                raise Step0Error(f"Saved ECM search was not found: {search_name}")
        try:
            await search_panel_locator.first.wait_for(state="visible", timeout=timeout_ms)
        except Exception as error:
            raise Step0Error("Saved ECM search did not open.") from error
        visible_panel = await _visible_locators(search_panel_locator)
        if len(visible_panel) != 1:
            raise Step0Error("Saved ECM search panel is not unique.")
    await _ensure_lookup_criteria_expanded(visible_panel[0], timeout_ms)
    await _wait_for_lookup_criteria_hydration(visible_panel[0], timeout_ms)
    return visible_panel[0]


async def _wait_ready(page, timeout_ms: int) -> None:
    deadline = asyncio.get_running_loop().time() + timeout_ms / 1000
    while asyncio.get_running_loop().time() < deadline:
        try:
            if await page.get_by_role("button", name="Search", exact=True).count():
                return
        except Exception:
            pass
        await asyncio.sleep(0.25)
    raise Step0Error("ECM Navigator did not become ready before the timeout.")


async def _select_ecm_page(context, ecm_url: str):
    target = urlparse(ecm_url)
    for page in context.pages:
        current = urlparse(str(getattr(page, "url", "")))
        if (current.scheme, current.netloc) == (target.scheme, target.netloc):
            return page
    return context.pages[0] if context.pages else await context.new_page()


async def _ecm_page_state(page, ecm_url: str, context) -> dict[str, object]:
    try:
        document = await page.evaluate("""() => ({
            ready_state: document.readyState,
            body_text_length: (document.body?.innerText || '').trim().length,
            body_child_count: document.body?.childElementCount || 0
        })""")
    except Exception:
        document = {"ready_state": "<unavailable>", "body_text_length": 0, "body_child_count": 0}
    selectors = ("#icm_widget_SelectorTabContainer_0_tablist_menuBtn", "[id$='_tablist_menuBtn']", "[role='tablist']", "[role='main']", "[widgetid]")
    counts = {}
    for selector in selectors:
        try:
            counts[selector] = await page.locator(selector).count()
        except Exception:
            counts[selector] = 0
    current, expected = urlparse(str(getattr(page, "url", ""))), urlparse(ecm_url)
    return {**document, "url": _sanitise_url(getattr(page, "url", "")),
            "origin_matches": (current.scheme, current.netloc) == (expected.scheme, expected.netloc),
            "page_count": len(context.pages), "shell_counts": counts}


async def _wait_for_ecm_readiness(page, ecm_url: str, context, timeout_ms: int) -> dict[str, object]:
    loop = asyncio.get_running_loop()
    started = loop.time()
    deadline = started + timeout_ms / 1000
    reload_at = started + max(2.0, min(15.0, timeout_ms / 2000))
    reloaded = False
    state: dict[str, object] = {}
    while loop.time() < deadline:
        if callable(getattr(page, "is_closed", None)) and page.is_closed():
            raise Step0Error("The ECM browser page was closed.")
        state = await _ecm_page_state(page, ecm_url, context)
        if await _page_requires_authentication(page):
            raise AuthenticationRequired("ECM authentication is required before Navigator can load.")
        shell_counts = state.get("shell_counts", {})
        if (state.get("origin_matches") and state.get("ready_state") in {"interactive", "complete"}
                and (state.get("body_child_count", 0) or state.get("body_text_length", 0))
                and any(value > 0 for value in shell_counts.values())):
            _log("page_readiness_ready", **state)
            return state
        # Never reload an identity provider, an unknown origin, or an active login.
        if not reloaded and state.get("origin_matches") and loop.time() >= reload_at:
            reloaded = True
            _log("page_readiness_reload", **state)
            try:
                await page.reload(wait_until="domcontentloaded", timeout=max(1, int((deadline - loop.time()) * 1000)))
            except Exception as error:
                _log("page_readiness_reload_failed", error=type(error).__name__)
        await asyncio.sleep(min(0.5, max(0, deadline - loop.time())))
    reason = ("wrong_origin" if not state.get("origin_matches") else
              "missing_ecm_shell" if state.get("body_child_count") or state.get("body_text_length") else "empty_dom")
    _log("page_readiness_failed", reason=reason, **state)
    raise PageNotReadyError(reason, state)


async def _first_visible(candidates):
    for locator in candidates:
        visible = await _visible_locators(locator)
        if visible:
            return visible[0]
    return None


async def _wait_for_search_panel(page, timeout_ms: int):
    panel = page.get_by_role("tabpanel", name="Search", exact=True)
    deadline = asyncio.get_running_loop().time() + timeout_ms / 1000
    for attempt in range(3):
        if await _visible_locators(panel):
            return panel
        if await _page_requires_authentication(page):
            raise AuthenticationRequired("ECM authentication is required before the Search panel can open.")
        menu = await _first_visible([page.locator("#icm_widget_SelectorTabContainer_0_tablist_menuBtn"),
                                     page.locator("[id$='_tablist_menuBtn']"),
                                     page.get_by_role("button", name=re.compile("menu", re.IGNORECASE))])
        if menu is not None:
            await click_control(menu, min(timeout_ms, 10_000))
            await asyncio.sleep(min(0.75, max(0, deadline - asyncio.get_running_loop().time())))
        item = await _first_visible([page.get_by_role("menuitem", name="Search", exact=True),
                                     page.locator("[role='menuitem']").filter(has_text=re.compile(r"^Search$")),
                                     page.locator("#dijit_MenuItem_6")])
        if item is not None:
            try:
                await click_control(item, min(timeout_ms, 10_000))
                await panel.first.wait_for(state="visible", timeout=max(1, min(5000, int((deadline - asyncio.get_running_loop().time()) * 1000))))
                return panel
            except Exception:
                pass
        if asyncio.get_running_loop().time() >= deadline:
            break
        await asyncio.sleep(min(0.5 * (attempt + 1), max(0, deadline - asyncio.get_running_loop().time())))
    if await _page_requires_authentication(page):
        raise AuthenticationRequired("ECM authentication is required before the Search panel can open.")
    raise Step0Error("ECM Search panel did not open after selecting Search.")


async def _expand_treeitem(item, timeout_ms: int) -> None:
    if await item.get_attribute("aria-expanded") == "true":
        return
    node = item.locator("xpath=ancestor::div[contains(@class, 'dijitTreeNode')][1]")
    expando = await _visible_locators(node.locator(".dijitTreeExpando"))
    await click_control(expando[0] if expando else item, timeout_ms)
    deadline = asyncio.get_running_loop().time() + timeout_ms / 1000
    while asyncio.get_running_loop().time() < deadline:
        if await item.get_attribute("aria-expanded") == "true" or await _visible_locators(node.locator(".dijitTreeNodeContainer [role='treeitem']")):
            return
        await asyncio.sleep(0.25)
    raise Step0Error("ECM search tree section did not expand before the timeout.")


async def _saved_searches_in_group(page, group_name: str, search_name: str):
    labels = await _visible_locators(page.locator(".dijitTreeLabel[role='treeitem']").filter(has_text=group_name))
    if len(labels) > 1:
        raise Step0Error(f"ECM search section is ambiguous: {group_name}")
    if labels:
        node = labels[0].locator("xpath=ancestor::div[contains(@class, 'dijitTreeNode')][1]")
        return await _visible_locators(node.locator(".dijitTreeNodeContainer [role='treeitem']").filter(has_text=re.compile(r"^" + re.escape(search_name) + r"$")))
    groups = await _visible_locators(page.get_by_role("group", name=group_name, exact=True))
    if len(groups) > 1:
        raise Step0Error(f"ECM search section is ambiguous: {group_name}")
    return await _visible_locators(groups[0].get_by_role("treeitem", name=search_name, exact=True)) if groups else []


async def _criteria_hydration_state(form) -> dict[str, object]:
    if not callable(getattr(form, "evaluate", None)):
        return {"probe_supported": False}
    return await form.evaluate("""element => {
        const expected = ['ECM_DocDate', 'ECM_OCRStatus', 'ECM_Title', 'ECM_InOut'];
        const area = element.querySelector('.ecmSearchFormInputArea');
        const rows = area ? Array.from(area.querySelectorAll('tr.ecmSearchFormFieldArea')) : [];
        const properties = Array.from(element.querySelectorAll('[name^="ECM_"]')).map(n => n.getAttribute('name'));
        return {probe_supported: true, input_area_present: Boolean(area), row_count: rows.length,
            visible_row_count: rows.filter(row => {
                const style = window.getComputedStyle(row), box = row.getBoundingClientRect();
                return style.display !== 'none' && style.visibility !== 'hidden' && box.width > 0 && box.height > 0;
            }).length, missing_properties: expected.filter(p => !properties.includes(p))};
    }""")


async def _wait_for_criteria_hydration(panel, timeout_ms: int) -> None:
    deadline = asyncio.get_running_loop().time() + timeout_ms / 1000
    state: dict[str, object] = {}
    while asyncio.get_running_loop().time() < deadline:
        try:
            state = await _criteria_hydration_state(await _active_form(panel))
        except Exception as error:
            state = {"probe_supported": True, "probe_error": type(error).__name__}
        if not state.get("probe_supported", True) or (state.get("input_area_present") and not state.get("missing_properties") and state.get("row_count", 0) >= 4 and state.get("visible_row_count", 0) >= 4):
            return
        await asyncio.sleep(0.25)
    _log("criteria_hydration_failed", state=state)
    raise CriteriaHydrationError(state)


async def _open_saved_search(page, timeout_ms: int, search_panel=None, *, search_name: str = ECM_SEARCH_NAME):
    # Lookup uses a different search and retains its existing shared-helper semantics.
    if search_name != ECM_SEARCH_NAME:
        return await _open_lookup_search(page, timeout_ms, search_panel, search_name=search_name)
    panel = page.get_by_role("tabpanel", name=search_name, exact=True)
    visible = await _visible_locators(panel)
    if len(visible) > 1:
        raise Step0Error("Saved ECM search panel is ambiguous.")
    if not visible:
        if search_panel is None:
            search_panel = await _wait_for_search_panel(page, timeout_ms)
        deadline = asyncio.get_running_loop().time() + timeout_ms / 1000
        groups = {}
        while asyncio.get_running_loop().time() < deadline:
            for name in ("Recent Searches", "All Searches"):
                groups[name] = await _visible_locators(page.get_by_role("treeitem", name=name, exact=True))
            if any(groups.values()):
                break
            if await _visible_locators(panel):
                break
            await asyncio.sleep(0.25)
        if not await _visible_locators(panel):
            for name in ("Recent Searches", "All Searches"):
                matches = groups.get(name, [])
                if len(matches) > 1:
                    raise Step0Error(f"ECM search section is ambiguous: {name}")
                if not matches:
                    continue
                await _expand_treeitem(matches[0], timeout_ms)
                saved = await _saved_searches_in_group(page, name, search_name)
                if len(saved) > 1:
                    raise Step0Error(f"Saved ECM search is ambiguous in {name}.")
                if saved:
                    await click_control(saved[0], timeout_ms)
                    break
            else:
                raise Step0Error(f"Saved ECM search was not found: {search_name}")
        try:
            await panel.first.wait_for(state="visible", timeout=timeout_ms)
        except Exception as error:
            raise Step0Error(f"Saved ECM search did not open: {search_name}") from error
        visible = await _visible_locators(panel)
        if len(visible) != 1:
            raise Step0Error("Saved ECM search panel is not unique.")
    await _wait_for_criteria_hydration(visible[0], timeout_ms)
    return visible[0]


async def _page_requires_authentication(page) -> bool:
    try:
        url = str(page.url).casefold()
    except Exception:
        url = ""
    if any(marker in url for marker in ("/login", "signin", "sso", "authenticate", "signon")):
        return True
    try:
        title = (await page.title()).casefold()
        body = (await page.locator("body").inner_text()).casefold()
    except Exception:
        return False
    return any(marker in title or marker in body[:5000] for marker in ("sign in", "signin", "log in", "single sign-on", "authenticate"))


async def _search_period(page, panel, start_date: date, end_date: date, timeout_ms: int) -> int:
    form = await _active_form(panel)
    for label, value in (("OCR Status", "Processing"), ("Document title:", "Method Statement"), ("In/Out", "OUT")):
        await _assert_value(form, label, value)
    operator_matches = []
    for control in await _labeled_controls(form, "Document date", include_operator=True):
        metadata = await _control_metadata(control)
        value = (await _step0_control_value(control)).casefold()
        if "Operator" in metadata.get("id", "") and value == "between":
            operator_matches.append(control)
    if len(operator_matches) != 1:
        raise Step0Error("Document date operator 'Between' was not found uniquely.")
    date_inputs = await _visible_locators(form.locator('input[placeholder*="M/d/yyyy"]'))
    if len(date_inputs) != 2:
        date_inputs = []
        for control in await _visible_locators(form.locator("input")):
            metadata = await _control_metadata(control)
            text = " ".join(metadata.get(key, "") for key in ("aria", "name", "placeholder", "context")).casefold()
            if "document date" in text or "m/d/yyyy" in text:
                date_inputs.append(control)
    if len(date_inputs) != 2:
        raise Step0Error("Expected exactly two Document date inputs.")
    for control, value in zip(date_inputs, (_date_input_value(start_date), _date_input_value(end_date)), strict=True):
        await control.fill(value)
        await control.dispatch_event("change")
    await date_inputs[1].press("Tab")
    search = await _visible_locators(panel.get_by_role("button", name="Search", exact=True))
    if len(search) != 1:
        raise Step0Error("ECM Search button was not found uniquely.")
    await click_control(search[0], timeout_ms)
    toolbar = panel.get_by_role("toolbar", name="SearchResultsToolbar")
    await toolbar.first.wait_for(state="visible", timeout=timeout_ms)
    body = await page.locator("body").inner_text()
    count = re.search(r"The search returned\s+([\d,]+)\s+items?", body, re.IGNORECASE)
    if count:
        result_count = int(count.group(1).replace(",", ""))
    else:
        rows = await panel.get_by_role("grid", name="grid").get_by_role("row").count()
        result_count = max(rows - 1, 0)
    if not result_count:
        raise NoSearchResults(f"ECM search returned no documents for {start_date} to {end_date}.")
    return result_count


async def _export_csv(page, panel, temporary_path: Path, timeout_ms: int) -> None:
    buttons = await _visible_locators(panel.get_by_role("button", name="Export All", exact=True))
    if len(buttons) != 1:
        buttons = await _visible_locators(page.get_by_role("button", name="Export All", exact=True))
    if len(buttons) != 1:
        raise Step0Error("ECM Export All button was not found uniquely.")
    await click_control(buttons[0], timeout_ms)
    dialog = page.get_by_role("dialog", name="Export All", exact=True)
    await dialog.first.wait_for(state="visible", timeout=min(timeout_ms, 10_000))
    dialogs = await _visible_locators(dialog)
    if len(dialogs) != 1:
        raise Step0Error("ECM Export All dialog is ambiguous.")
    dialog = dialogs[0]
    all_items = await _visible_locators(dialog.get_by_role("radio", name="Export all items", exact=True))
    raw_values = await _visible_locators(dialog.get_by_role("checkbox", name="Export raw values", exact=True))
    if len(all_items) != 1 or len(raw_values) != 1:
        raise Step0Error("Export All dialog options are missing or ambiguous.")
    if not await all_items[0].is_checked():
        await click_control(all_items[0], timeout_ms)
        if not await all_items[0].is_checked():
            raise Step0Error("ECM Export all items could not be selected.")
    if await raw_values[0].is_checked():
        raise Step0Error("Export raw values must remain unchecked.")
    export = await _visible_locators(dialog.get_by_role("button", name="Export", exact=True))
    if len(export) != 1:
        raise Step0Error("Export All confirmation was not found uniquely.")
    temporary_path.parent.mkdir(parents=True, exist_ok=True)
    async with page.expect_download(timeout=timeout_ms) as download_info:
        await click_control(export[0], timeout_ms)
    download = await download_info.value
    failure = await download.failure()
    if failure:
        raise Step0Error(f"ECM CSV download failed: {failure}")
    await download.save_as(str(temporary_path))


async def _wait_for_manual_authentication(page, timeout_seconds: float) -> None:
    """Wait only for the login screen to clear; shell readiness has its own timeout."""
    from .browser_support import focus_login_username

    _log("manual_authentication_required", timeout_seconds=timeout_seconds)
    deadline = asyncio.get_running_loop().time() + timeout_seconds
    focused_login = False
    while asyncio.get_running_loop().time() < deadline:
        if callable(getattr(page, "is_closed", None)) and page.is_closed():
            raise AuthenticationRequired("The ECM browser page was closed during authentication.")
        authentication_required = await _page_requires_authentication(page)
        if not authentication_required:
            _log("authentication_detected")
            return
        if not focused_login:
            await focus_login_username(page)
            focused_login = True
        await asyncio.sleep(min(1, max(0, deadline - asyncio.get_running_loop().time())))
    raise AuthenticationRequired(f"Manual ECM SSO/MFA authentication timed out after {timeout_seconds} seconds.")


async def run_export(*, start_date: date, end_date: date, output_path: Path, period_input_path: Path | None, ecm_url: str, profile_dir: Path, timeout_seconds: int, root_dir: Path, login_if_needed: bool = True, auth_timeout_seconds: int = 120, diagnostics_dir: Path | None = None, reset_profile_requested: bool = False, headed: bool = True, failure_browser_hold_seconds: int = 120) -> int:
    root = V3_ROOT.resolve()
    if root_dir.resolve() != root:
        raise Step0Error("The Step 0 worker root is fixed to the V3 control-panel directory.")
    for path, label in ((output_path, "output"), (profile_dir, "Edge profile")):
        if not _inside(path, root):
            raise Step0Error(f"V3 {label} path escapes the V3 root: {path}")
    if profile_dir.resolve() == root:
        raise Step0Error("The V3 Edge profile must be a directory below the V3 root.")
    if period_input_path and not _inside(period_input_path, root):
        raise Step0Error(f"V3 period CSV path escapes the V3 root: {period_input_path}")
    if diagnostics_dir and not _inside(diagnostics_dir, root):
        raise Step0Error(f"V3 diagnostics path escapes the V3 root: {diagnostics_dir}")
    if period_input_path and output_path.resolve() == period_input_path.resolve():
        raise Step0Error("The run output and period CSV must be different paths.")
    if timeout_seconds <= 0 or auth_timeout_seconds <= 0 or failure_browser_hold_seconds < 0:
        raise Step0Error("Timeouts must be positive and the failure browser hold must be non-negative.")
    if end_date < start_date:
        raise Step0Error("End date cannot be earlier than start date.")
    if output_path.exists():
        raise Step0Error(f"Refusing to overwrite existing run output: {output_path}")
    if reset_profile_requested:
        backup = reset_profile(profile_dir)
        _log("profile_reset", backup_profile=backup)
    profile_dir.mkdir(parents=True, exist_ok=True)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary = output_path.with_name(f".{output_path.name}.download")
    if temporary.exists():
        temporary.unlink()
    try:
        from playwright.async_api import async_playwright
    except ImportError as error:
        raise Step0Error("Playwright is required in the V3 environment.") from error
    context = None
    page = None
    browser_diagnostics = BrowserDiagnostics()
    stage = "starting ECM browser"
    visible = headed
    headless_fallback_used = False
    auth_deadline = None
    resume = False
    try:
        async with async_playwright() as playwright:
            try:
                while True:
                    if context is None:
                        context = await playwright.chromium.launch_persistent_context(user_data_dir=str(profile_dir), channel="msedge", headless=not visible, accept_downloads=True)
                        page = await _select_ecm_page(context, ecm_url)
                        browser_diagnostics.attach(page)
                    try:
                        if not resume:
                            stage = "opening ECM Navigator"
                            try:
                                await page.goto(ecm_url, wait_until="domcontentloaded", timeout=timeout_seconds * 1000)
                            except Exception:
                                if await _page_requires_authentication(page):
                                    raise AuthenticationRequired("ECM authentication is required before Navigator can load.")
                                raise
                        resume = False
                        stage = "waiting for ECM Navigator readiness"
                        await _wait_for_ecm_readiness(page, ecm_url, context, timeout_seconds * 1000)
                        stage = "opening the Search page"
                        search_panel = await _wait_for_search_panel(page, timeout_seconds * 1000)
                        stage = "opening the saved search"
                        panel = await _open_saved_search(page, timeout_seconds * 1000, search_panel)
                        stage = "running the saved search"
                        count = await _search_period(page, panel, start_date, end_date, timeout_seconds * 1000)
                        stage = "downloading Export All"
                        await _export_csv(page, panel, temporary, timeout_seconds * 1000)
                        break
                    except AuthenticationRequired:
                        if not login_if_needed:
                            raise
                        if not visible:
                            if headless_fallback_used:
                                raise AuthenticationRequired("ECM authentication is still required after the visible login attempt.")
                            headless_fallback_used = True
                            await context.close()
                            context = None
                            visible = True
                            continue
                        stage = "waiting for manual authentication"
                        if auth_deadline is None:
                            auth_deadline = asyncio.get_running_loop().time() + auth_timeout_seconds
                        remaining = auth_deadline - asyncio.get_running_loop().time()
                        if remaining <= 0:
                            raise AuthenticationRequired(f"Manual ECM SSO/MFA authentication timed out after {auth_timeout_seconds} seconds.")
                        await _wait_for_manual_authentication(page, remaining)
                        if not headed:
                            # Wait for the authenticated shell before reopening the persistent profile.
                            stage = "waiting for ECM Navigator readiness"
                            await _wait_for_ecm_readiness(page, ecm_url, context, timeout_seconds * 1000)
                            await context.close()
                            context = None
                            visible = False
                            _log("resuming_headless_export")
                        else:
                            # Continue in the same page without interrupting the SSO redirect.
                            resume = True
                stage = "validating and publishing CSV"
                _publish_validated_exports(temporary, output_path, period_input_path)
            except Exception as error:
                _log("stage_failed", stage=stage, error=type(error).__name__, url=_sanitise_url(getattr(page, "url", "")))
                if diagnostics_dir and page is not None:
                    try:
                        await save_failure_diagnostics(page, diagnostics_dir, "step0", browser_diagnostics)
                    except Exception as diagnostic_error:
                        _log("failure_diagnostics_failed", error=type(diagnostic_error).__name__)
                if context is not None and visible and failure_browser_hold_seconds > 0:
                    _log("browser_retained_for_inspection", seconds=failure_browser_hold_seconds)
                    await asyncio.sleep(failure_browser_hold_seconds)
                raise
            finally:
                if context is not None:
                    await context.close()
                    context = None
        _log("workflow_complete", result_count=count, output=output_path)
        return count
    finally:
        temporary.unlink(missing_ok=True)


def _parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError as error:
        raise argparse.ArgumentTypeError("Use YYYY-MM-DD.") from error


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Export the V3 ECM SMM_ECM_Reporting saved search.")
    parser.add_argument("--start-date", required=True, type=_parse_date)
    parser.add_argument("--end-date", required=True, type=_parse_date)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--period-input", type=Path)
    parser.add_argument("--ecm-url", default=os.getenv("CONTROL_PANEL_V3_ECM_URL", DEFAULT_ECM_URL))
    parser.add_argument("--profile-dir", required=True, type=Path)
    parser.add_argument("--root-dir", type=Path, default=V3_ROOT, help=argparse.SUPPRESS)
    parser.add_argument("--timeout-seconds", type=int, default=60)
    parser.add_argument("--headed", action=argparse.BooleanOptionalAction, default=True)
    parser.add_argument("--login-if-needed", action=argparse.BooleanOptionalAction, default=True)
    parser.add_argument("--auth-timeout-seconds", type=int, default=120)
    parser.add_argument("--failure-browser-hold-seconds", type=int, default=120)
    parser.add_argument("--diagnostics-dir", type=Path)
    parser.add_argument("--reset-profile", action="store_true")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.timeout_seconds <= 0 or args.auth_timeout_seconds <= 0 or args.failure_browser_hold_seconds < 0:
        raise Step0Error("Timeouts must be positive and the failure browser hold must be non-negative.")
    result = asyncio.run(run_export(start_date=args.start_date, end_date=args.end_date, output_path=args.output, period_input_path=args.period_input, ecm_url=args.ecm_url, profile_dir=args.profile_dir, timeout_seconds=args.timeout_seconds, root_dir=args.root_dir, login_if_needed=args.login_if_needed, auth_timeout_seconds=args.auth_timeout_seconds, diagnostics_dir=args.diagnostics_dir, reset_profile_requested=args.reset_profile, headed=args.headed, failure_browser_hold_seconds=args.failure_browser_hold_seconds))
    print(f"Search: {ECM_SEARCH_NAME}\nSearch results: {result}\nOutput CSV: {args.output.resolve()}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (Step0Error, CsvValidationError, OSError, ValueError) as error:
        print(f"Error: {error}", file=sys.stderr)
        raise SystemExit(1)
