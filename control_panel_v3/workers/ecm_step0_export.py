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
}
CRITERION_VALUE_ALIASES = {"ECM_OCRStatus": {"0": "Not processed", "1": "Queued", "2": "Processing"}, "ECM_InOut": {"0": "IN", "1": "BOTH", "2": "OUT"}}
DEFAULT_ECM_URL = "https://ecm.chunwo.com/navigator/"
V3_ROOT = Path(__file__).resolve().parents[1]


class Step0Error(RuntimeError):
    """The ECM export could not safely complete."""


class NoSearchResults(Step0Error):
    """The saved search returned no exportable records."""


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
                labels: labels.join(' | '), properties: row ? Array.from(row.querySelectorAll('[name^="ECM_"]')).map(node => node.getAttribute('name') || '').join(' | ') : '',
                context: ancestors.join(' | ')
            };
        }"""
    )


async def _labeled_controls(form, label: str, *, include_select: bool = True, include_operator: bool = False) -> list:
    selector = "input,textarea,[role='textbox'],[role='combobox']" + (",select" if include_select else "")
    controls = form.locator(selector)
    property_name = CRITERION_PROPERTIES.get(label)
    candidates = []
    if property_name and label != "Document date":
        direct = form.locator(
            f'[name="{property_name}"]:not([type="hidden"]),[title="{label.rstrip(":")}"],[alt="{label.rstrip(":")}"]'
        )
        candidates = await _visible_locators(direct)
        if not candidates:
            rows = form.locator(f'tr:has([name="{property_name}"]), .ecmSearchFormFieldArea:has([name="{property_name}"])')
            for index in range(await rows.count()):
                candidates.extend(await _visible_locators(rows.nth(index).locator(selector)))
    property_candidates_found = bool(candidates)
    if not candidates:
        candidates = [controls.nth(index) for index in range(await controls.count())]
    wanted = label.casefold().rstrip(":")
    result = []
    seen: set[str] = set()
    for index, control in enumerate(candidates):
        if not await control.is_visible() or await control.get_attribute("type") == "hidden":
            continue
        metadata = await _control_metadata(control)
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


async def _control_value(control) -> str:
    for method in ("input_value", "inner_text"):
        try:
            value = await getattr(control, method)()
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
                    widget?.innerText, element.parentElement?.innerText, property?.value]
                    .find(value => String(value || '').trim()) || '';
            }"""
        )
        return str(value).strip()
    except Exception:
        return ""


async def _criterion_value(control, label: str) -> str:
    value = await _control_value(control)
    aliases = CRITERION_VALUE_ALIASES.get(CRITERION_PROPERTIES.get(label, ""), {})
    return aliases.get(value.casefold(), value)


async def _assert_value(form, label: str, expected: str) -> None:
    controls = await _labeled_controls(form, label)
    values = [await _criterion_value(control, label) for control in controls]
    matches = [value for value in values if expected.casefold() in value.casefold()]
    _log("criteria_values", criterion=label, values=values)
    if len(matches) != 1:
        raise Step0Error(f"Saved search criterion {label!r} must be exactly {expected!r}; observed {values!r}.")


async def _active_form(search_panel):
    for selector in ("[role='search']", "#ecm_widget_search_SearchForm_0"):
        visible = await _visible_locators(search_panel.locator(selector))
        if len(visible) == 1:
            return visible[0]
    return search_panel


async def _open_saved_search(page, timeout_ms: int, search_panel=None, *, search_name: str = ECM_SEARCH_NAME):
    search_panel_locator = page.get_by_role("tabpanel", name=search_name, exact=True)
    visible_panel = await _visible_locators(search_panel_locator)
    if len(visible_panel) == 1:
        return visible_panel[0]
    _log("saved_search_navigation", search=search_name)
    all_searches = page.get_by_role("treeitem", name="All Searches", exact=True)
    visible = await _visible_locators(all_searches)
    if len(visible) != 1:
        raise Step0Error("ECM All Searches navigation was not found uniquely.")
    expanded = await visible[0].get_attribute("aria-expanded")
    if expanded != "true":
        await click_control(visible[0], timeout_ms)
    saved = page.get_by_role("treeitem", name=search_name, exact=True)
    matches = await _visible_locators(saved)
    if len(matches) != 1:
        group = page.get_by_role("treeitem", name="All Searches", exact=True)
        if len(await _visible_locators(group)) == 1:
            await click_control(group.get_by_role("button").first, timeout_ms)
        matches = await _visible_locators(saved)
    if len(matches) != 1:
        raise Step0Error(f"Saved ECM search was not found uniquely: {search_name}")
    await click_control(matches[0], timeout_ms)
    try:
        await search_panel_locator.first.wait_for(state="visible", timeout=timeout_ms)
    except Exception as error:
        raise Step0Error("Saved ECM search did not open.") from error
    panel = search_panel_locator.first
    if hasattr(panel, "get_by_role"):
        try:
            await panel.get_by_role("button", name="Search", exact=True).wait_for(
                state="visible", timeout=timeout_ms
            )
        except Exception as error:
            raise Step0Error("ECM saved-search criteria did not finish loading.") from error
    return panel


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


async def _page_requires_authentication(page) -> bool:
    try:
        url = str(page.url).casefold()
    except Exception:
        url = ""
    if any(marker in url for marker in ("/login", "sso", "authenticate", "signon")):
        return True
    try:
        title = (await page.title()).casefold()
        body = (await page.locator("body").inner_text()).casefold()
    except Exception:
        return False
    return any(marker in title or marker in body[:5000] for marker in ("sign in", "log in", "single sign-on", "authenticate"))


async def _search_period(page, panel, start_date: date, end_date: date, timeout_ms: int) -> int:
    form = await _active_form(panel)
    for label, value in (("OCR Status", "Processing"), ("Document title:", "Method Statement"), ("In/Out", "OUT")):
        await _assert_value(form, label, value)
    operator_matches = []
    for control in await _labeled_controls(form, "Document date", include_operator=True):
        metadata = await _control_metadata(control)
        value = (await _control_value(control)).casefold()
        if "Operator" in metadata.get("id", "") and value == "between":
            operator_matches.append(control)
    if len(operator_matches) != 1:
        raise Step0Error("Document date operator 'Between' was not found uniquely.")
    date_inputs = await _visible_locators(form.locator('input[placeholder*="M/d/yyyy"]'))
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
    all_items = await _visible_locators(dialog.get_by_role("radio", name="Export all items", exact=True))
    raw_values = await _visible_locators(dialog.get_by_role("checkbox", name="Export raw values", exact=True))
    if len(all_items) != 1 or len(raw_values) != 1:
        raise Step0Error("Export All dialog options are missing or ambiguous.")
    if not await all_items[0].is_checked():
        await all_items[0].check()
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


async def _authenticated_page(playwright, url: str, profile: Path, timeout_seconds: int, diagnostics: BrowserDiagnostics | None = None):
    _log("manual_authentication_required", timeout_seconds=timeout_seconds)
    context = await playwright.chromium.launch_persistent_context(user_data_dir=str(profile), channel="msedge", headless=False, accept_downloads=True)
    try:
        page = context.pages[0] if context.pages else await context.new_page()
        if diagnostics:
            diagnostics.attach(page)
        await page.goto(url, wait_until="domcontentloaded", timeout=timeout_seconds * 1000)
        deadline = asyncio.get_running_loop().time() + timeout_seconds
        while asyncio.get_running_loop().time() < deadline:
            try:
                if await _page_requires_authentication(page):
                    await asyncio.sleep(1)
                    continue
                await _wait_ready(page, 2000)
                _log("authentication_ready")
                return
            except Exception:
                await asyncio.sleep(1)
        raise Step0Error("Manual ECM SSO/MFA authentication timed out.")
    finally:
        await context.close()


async def run_export(*, start_date: date, end_date: date, output_path: Path, period_input_path: Path | None, ecm_url: str, profile_dir: Path, timeout_seconds: int, root_dir: Path, login_if_needed: bool = False, auth_timeout_seconds: int = 120, diagnostics_dir: Path | None = None, reset_profile_requested: bool = False) -> int:
    root = V3_ROOT.resolve()
    if root_dir.resolve() != root:
        raise Step0Error("The Step 0 worker root is fixed to the V3 control-panel directory.")
    for path, label in ((output_path, "output"), (profile_dir, "Edge profile")):
        if not _inside(path, root):
            raise Step0Error(f"V3 {label} path escapes the V3 root: {path}")
    if period_input_path and not _inside(period_input_path, root):
        raise Step0Error(f"V3 period CSV path escapes the V3 root: {period_input_path}")
    if diagnostics_dir and not _inside(diagnostics_dir, root):
        raise Step0Error(f"V3 diagnostics path escapes the V3 root: {diagnostics_dir}")
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
    browser_diagnostics = BrowserDiagnostics()
    try:
        async with async_playwright() as playwright:
            while True:
                context = await playwright.chromium.launch_persistent_context(user_data_dir=str(profile_dir), channel="msedge", headless=True, accept_downloads=True)
                try:
                    page = context.pages[0] if context.pages else await context.new_page()
                    browser_diagnostics.attach(page)
                    await page.goto(ecm_url, wait_until="domcontentloaded", timeout=timeout_seconds * 1000)
                    if await _page_requires_authentication(page):
                        raise Step0Error("ECM authentication is required in the V3 dedicated Edge profile.")
                    await _wait_ready(page, timeout_seconds * 1000)
                    panel = await _open_saved_search(page, timeout_seconds * 1000)
                    count = await _search_period(page, panel, start_date, end_date, timeout_seconds * 1000)
                    await _export_csv(page, panel, temporary, timeout_seconds * 1000)
                    break
                except Exception as error:
                    _log("stage_failed", error=type(error).__name__, message=str(error)[:500], url=_sanitise_url(getattr(locals().get("page"), "url", "")))
                    if login_if_needed and auth_timeout_seconds > 0 and "page" in locals():
                        page = locals()["page"]
                        if await _page_requires_authentication(page):
                            await context.close()
                            context = None
                            await _authenticated_page(playwright, ecm_url, profile_dir, auth_timeout_seconds, browser_diagnostics)
                            login_if_needed = False
                            continue
                    raise
                finally:
                    if context is not None:
                        await context.close()
                        context = None
        _publish_validated_exports(temporary, output_path, period_input_path)
        _log("workflow_complete", result_count=count, output=output_path)
        return count
    except Exception:
        page = locals().get("page")
        if diagnostics_dir and page is not None:
            await save_failure_diagnostics(page, diagnostics_dir, "step0", browser_diagnostics)
        raise
    finally:
        if context is not None:
            await context.close()
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
    parser.add_argument("--timeout-seconds", type=int, default=120)
    parser.add_argument("--login-if-needed", action="store_true")
    parser.add_argument("--auth-timeout-seconds", type=int, default=120)
    parser.add_argument("--diagnostics-dir", type=Path)
    parser.add_argument("--reset-profile", action="store_true")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.timeout_seconds <= 0 or args.auth_timeout_seconds <= 0:
        raise Step0Error("Timeouts must be greater than zero.")
    result = asyncio.run(run_export(start_date=args.start_date, end_date=args.end_date, output_path=args.output, period_input_path=args.period_input, ecm_url=args.ecm_url, profile_dir=args.profile_dir, timeout_seconds=args.timeout_seconds, root_dir=args.root_dir, login_if_needed=args.login_if_needed, auth_timeout_seconds=args.auth_timeout_seconds, diagnostics_dir=args.diagnostics_dir, reset_profile_requested=args.reset_profile))
    print(f"Search: {ECM_SEARCH_NAME}\nSearch results: {result}\nOutput CSV: {args.output.resolve()}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (Step0Error, CsvValidationError, OSError, ValueError) as error:
        print(f"Error: {error}", file=sys.stderr)
        raise SystemExit(1)
