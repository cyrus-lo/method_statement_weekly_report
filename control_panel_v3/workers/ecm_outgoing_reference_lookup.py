"""Find oldest ECM dates for eligible rows in a V3-owned review workbook.

A no-match is a row-level outcome, not a batch failure. Each successful date is
committed independently using an atomic workbook replacement.
"""

from __future__ import annotations

import argparse
import asyncio
import csv
import json
import os
import re
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

from .browser_support import BrowserDiagnostics, click_control, reset_profile, save_failure_diagnostics, wait_for_manual_auth
from .ecm_step0_export import AuthenticationRequired, CRITERION_VALUE_ALIASES, DEFAULT_ECM_URL, _active_form, _control_metadata, _control_value, _criterion_value, _ensure_lookup_criteria_expanded, _labeled_controls, _open_saved_search, _page_requires_authentication, _visible_locators
from .properties_csv import CSV_HEADERS, parse_ecm_datetime
from .reference_normalization import outgoing_reference_without_revision

LOOKUP_SEARCH_NAME = "Search ECM Document"

REMARK_COLUMN = 13
COUNT_COLUMN = 9
REVISION_COLUMN = 5
OUTGOING_REFERENCE_COLUMN = 16
OUTGOING_REFERENCE_HEADER = "Outgoing ref. no. (without revision)"
WORKBOOK_HEADERS = (
    "Project Code", "Document title", "Document date", "Name", "Document revision",
    "Modified By", "Major Version", "SMM ID", "Count (Y/N)", "Use SMM (Y/N)", "Use KMS (Y/N)",
)


class LookupError(RuntimeError):
    """Lookup could not safely search or update the workbook."""


def _log(event: str, **fields: object) -> None:
    print(json.dumps({"event": event, **fields}, ensure_ascii=True, default=str), flush=True)


def _inside(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root.resolve())
        return True
    except ValueError:
        return False


def _date_from_result(value: str) -> date:
    text = value.strip()
    if not text:
        raise LookupError("ECM result row has a blank Document date.")
    try:
        return date.fromisoformat(text)
    except ValueError:
        try:
            return parse_ecm_datetime(text, "Document date", 0).date()
        except ValueError as error:
            raise LookupError(f"ECM result has an unrecognized Document date: {text!r}.") from error


def _csv_date(value: str) -> str:
    return parse_ecm_datetime(value, "Document date", 0).date().isoformat()


def _identity(values: tuple[object, ...]) -> tuple[str, ...]:
    return tuple(str(value or "").strip().casefold() for value in values)


def read_outgoing_reference_map(csv_path: Path) -> dict[tuple[str, ...], str]:
    with csv_path.open("r", encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream)
        headers = tuple((header or "").strip() for header in (reader.fieldnames or ()))
        if headers != CSV_HEADERS:
            raise LookupError(f"Step 0 CSV headers are not the expected exact contract: {headers!r}.")
        mapping: dict[tuple[str, ...], str] = {}
        duplicates: set[tuple[str, ...]] = set()
        for row in reader:
            key = _identity((row["Document title"], _csv_date(row["Document date"]), row["Name"], row["Document revision"], row["Modified By"], row["Major Version"], row["SMM ID"]))
            reference = (row["Outgoing ref. no."] or "").strip()
            if key in mapping and mapping[key] != reference:
                duplicates.add(key)
            else:
                mapping[key] = reference
        if duplicates:
            raise LookupError(f"CSV contains ambiguous source identities ({len(duplicates)} row identities).")
        return mapping


def _atomic_save(workbook, workbook_path: Path, row_number: int) -> None:
    staged = workbook_path.with_name(f".{workbook_path.stem}.lookup-{row_number}.tmp{workbook_path.suffix}")
    try:
        workbook.save(staged)
        os.replace(staged, workbook_path)
    finally:
        staged.unlink(missing_ok=True)


def _write_summary(path: Path | None, *, worksheet: str, hits: int, misses: list[dict[str, object]], searched: int, scope: str) -> None:
    if path is None:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    staged = path.with_name(f".{path.name}.tmp")
    try:
        staged.write_text(json.dumps({
            "schema_version": 1,
            "worksheet": worksheet,
            "hits": hits,
            "misses": misses,
            "miss_count": len(misses),
            "searched": searched,
            "scope": scope,
            "partial_result": bool(misses),
        }, indent=2, ensure_ascii=True), encoding="utf-8")
        os.replace(staged, path)
    finally:
        staged.unlink(missing_ok=True)


def validate_workbook(worksheet) -> None:
    headings = tuple(str(worksheet.cell(1, column).value or "").strip() for column in range(1, 12))
    if headings != WORKBOOK_HEADERS:
        raise LookupError(f"Review workbook A:K headers do not match the expected contract: {headings!r}.")
    if str(worksheet.cell(1, REMARK_COLUMN).value or "").strip() != "Remark":
        raise LookupError("Review workbook column M must be headed 'Remark'.")
    if str(worksheet.cell(1, OUTGOING_REFERENCE_COLUMN).value or "").strip() != OUTGOING_REFERENCE_HEADER:
        raise LookupError(f"Review workbook column P must be headed {OUTGOING_REFERENCE_HEADER!r}.")


def build_search_rows(worksheet, reference_map: dict[tuple[str, ...], str], *, scope: str = "all", previously_unmatched: set[int] | None = None) -> list[tuple[int, str]]:
    if scope not in {"all", "blank", "previously-unmatched"}:
        raise LookupError(f"Unknown lookup rerun scope: {scope}")
    if scope == "previously-unmatched" and previously_unmatched is None:
        raise LookupError("Previously-unmatched scope requires a prior lookup summary.")
    tasks = []
    for row_number in range(2, worksheet.max_row + 1):
        if not any(worksheet.cell(row_number, column).value is not None for column in range(1, 9)):
            continue
        count_value = str(worksheet.cell(row_number, COUNT_COLUMN).value or "").strip().upper()
        revision = str(worksheet.cell(row_number, REVISION_COLUMN).value or "").strip()
        if count_value == "N" or not revision:
            continue
        source_identity = _identity((
            worksheet.cell(row_number, 2).value,
            _date_cell_iso(worksheet.cell(row_number, 3).value),
            worksheet.cell(row_number, 4).value,
            revision,
            worksheet.cell(row_number, 6).value,
            worksheet.cell(row_number, 7).value,
            worksheet.cell(row_number, 8).value,
        ))
        if source_identity not in reference_map:
            raise LookupError(f"Workbook row {row_number} cannot be matched to its exact Step 0 CSV record.")
        remark = str(worksheet.cell(row_number, REMARK_COLUMN).value or "").strip()
        if scope == "blank" and remark:
            continue
        if scope == "previously-unmatched" and (row_number not in (previously_unmatched or set()) or remark):
            continue
        reference = outgoing_reference_without_revision(
            str(worksheet.cell(row_number, OUTGOING_REFERENCE_COLUMN).value or ""), revision
        )
        if reference:
            tasks.append((row_number, reference))
        else:
            raise LookupError(f"Workbook row {row_number} has a blank Excel outgoing reference in column P.")
    return tasks


def _date_cell_iso(value: object) -> str:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value or "").strip()
    try:
        return date.fromisoformat(text).isoformat()
    except ValueError:
        return _csv_date(text)


async def _selected_in_out_value(control, fallback: str) -> str:
    """Read the selected ECM value, preferring the Dijit widget's hidden key."""
    try:
        raw_value = await control.evaluate("""element => {
            const widget = element.closest('.dijitComboBox') || element.closest('[role="combobox"]');
            const selected = widget?.querySelector('input[type="hidden"][name="ECM_InOut"]');
            return selected ? selected.value : null;
        }""")
    except Exception:
        raw_value = None
    if raw_value is None:
        return fallback
    value = str(raw_value).strip()
    aliases = CRITERION_VALUE_ALIASES.get("ECM_InOut", {})
    return aliases.get(value.casefold(), value)


def _is_operator_control(metadata: dict[str, str]) -> bool:
    class_name = metadata.get("class_name", "")
    return (
        "Operator" in metadata.get("id", "")
        or "Operator" in metadata.get("widget", "")
        or "operatorSelect" in class_name
        or "operatorsSelect" in class_name
    )


async def _click_visible_out_option(direction) -> bool:
    page = getattr(direction, "page", None)
    if page is None or not callable(getattr(page, "locator", None)):
        return False
    control_id = ""
    if callable(getattr(direction, "get_attribute", None)):
        try:
            control_id = (await direction.get_attribute("id")) or ""
        except Exception:
            control_id = ""
    selectors = []
    if control_id:
        selectors.append(f'[id="{control_id}_popup"] [role="option"]')
    selectors.append('[role="listbox"] [role="option"]')
    deadline = asyncio.get_running_loop().time() + 1.5
    while asyncio.get_running_loop().time() < deadline:
        for selector in selectors:
            for option in await _visible_locators(page.locator(selector)):
                try:
                    if (await option.inner_text()).strip().casefold() == "out":
                        await click_control(option, 2_000)
                        return True
                except Exception:
                    continue
        await asyncio.sleep(0.1)
    return False


async def _set_search_criteria(form, outgoing_reference: str) -> None:
    controls = await _labeled_controls(form, "In/Out")
    if len(controls) != 1:
        raise LookupError(f"ECM criterion 'In/Out' must have exactly one visible control; observed {len(controls)}.")
    direction = controls[0]
    current_value = await _criterion_value(direction, "In/Out")
    selected_value = await _selected_in_out_value(direction, current_value)
    if selected_value.casefold().strip() != "out":
        try:
            # Never press Enter inside ECMSearchForm: Enter submits the search
            # prematurely and collapses the Search Criteria section.
            await direction.fill("OUT")
            await asyncio.sleep(0.25)
            await direction.press("ArrowDown")
            if not await _click_visible_out_option(direction):
                await direction.dispatch_event("change")
            await direction.press("Tab")
        except Exception as error:
            raise LookupError("ECM criterion 'In/Out' could not be selected as 'OUT'.") from error
        deadline = asyncio.get_running_loop().time() + 5
        while asyncio.get_running_loop().time() < deadline:
            current_value = await _criterion_value(direction, "In/Out")
            selected_value = await _selected_in_out_value(direction, current_value)
            if selected_value.casefold().strip() == "out":
                break
            await asyncio.sleep(0.1)
        if selected_value.casefold().strip() != "out":
            page = getattr(direction, "page", None)
            if page is not None and callable(getattr(direction, "get_attribute", None)):
                try:
                    control_id = (await direction.get_attribute("id")) or ""
                    if control_id:
                        arrow = await _visible_locators(page.locator(f'[id="widget_{control_id}"] .dijitArrowButton'))
                        if len(arrow) == 1:
                            await click_control(arrow[0], 2_000)
                            if await _click_visible_out_option(direction):
                                await direction.press("Tab")
                                current_value = await _criterion_value(direction, "In/Out")
                                selected_value = await _selected_in_out_value(direction, current_value)
                except Exception:
                    pass
        if selected_value.casefold().strip() != "out":
            raise LookupError(f"ECM criterion 'In/Out' must be exactly 'OUT'; observed {selected_value!r} after selection.")
    _log("criteria_values", criterion="In/Out", values=[selected_value])
    ref_controls = await _labeled_controls(form, "Outgoing ref. no.", include_operator=True)
    operator_controls = []
    value_controls = []
    for control in ref_controls:
        metadata = await _control_metadata(control)
        if _is_operator_control(metadata):
            operator_controls.append(control)
        else:
            value_controls.append(control)
    operator_values = [await _control_value(control) for control in operator_controls]
    operator_matches = [
        control for control, value in zip(operator_controls, operator_values, strict=True)
        if value.casefold().strip() in {"starts with", "startswith"}
    ]
    if len(operator_matches) != 1:
        raise LookupError(
            "ECM Outgoing ref. no. operator must be exactly 'Starts With'; "
            f"observed {len(operator_controls)} operator control(s) with values {operator_values!r}."
        )
    if len(value_controls) != 1:
        raise LookupError(f"ECM Outgoing ref. no. value control is missing or ambiguous ({len(value_controls)} found).")
    await value_controls[0].fill(outgoing_reference)
    await value_controls[0].dispatch_event("change")
    await value_controls[0].press("Tab")
    observed_reference = await _control_value(value_controls[0])
    if observed_reference != outgoing_reference:
        raise LookupError(
            "ECM Outgoing ref. no. criterion did not retain the Excel value; "
            f"expected {outgoing_reference!r}, observed {observed_reference!r}."
        )
    _log("criteria_values", criterion="Outgoing ref. no.", operator="Starts With", values=[observed_reference])


async def _lookup_requires_authentication(page) -> bool:
    if await _page_requires_authentication(page):
        return True
    try:
        return bool(await _visible_locators(page.locator('.ecmLoginPane [type="password"]')))
    except Exception:
        return False


async def _open_lookup_with_auth(page, timeout_seconds: float):
    """Open the saved search, including when login appears after initial navigation."""
    deadline = asyncio.get_running_loop().time() + timeout_seconds
    while True:
        remaining = deadline - asyncio.get_running_loop().time()
        if remaining <= 0:
            raise LookupError("ECM saved search or authentication timed out.")
        if await _lookup_requires_authentication(page):
            _log("manual_authentication_required", timeout_seconds=remaining)
            await wait_for_manual_auth(page, _lookup_requires_authentication, remaining)
            _log("authentication_screen_cleared")
        try:
            remaining = deadline - asyncio.get_running_loop().time()
            if remaining <= 0:
                raise LookupError("ECM saved search or authentication timed out.")
            return await _open_saved_search(page, max(1, int(remaining * 1000)), search_name=LOOKUP_SEARCH_NAME)
        except Exception as error:
            if not isinstance(error, AuthenticationRequired) and not await _lookup_requires_authentication(page):
                raise LookupError("ECM saved search could not be opened or is not ready; inspect the V3 browser diagnostics.") from error
            remaining = deadline - asyncio.get_running_loop().time()
            if remaining <= 0:
                raise LookupError("ECM authentication timed out while opening the saved search.") from error
            _log("manual_authentication_required", timeout_seconds=remaining)
            await wait_for_manual_auth(page, _lookup_requires_authentication, remaining)
            _log("authentication_screen_cleared")
            await asyncio.sleep(min(0.1, max(0, deadline - asyncio.get_running_loop().time())))


async def _ensure_all_versions(panel, timeout_ms: int) -> None:
    """Apply All versions through the search-options dialog and verify its state."""
    buttons = await _visible_locators(panel.locator('.ecmSearchMoreOptions [role="button"][id$="_moreOptionsLink"]'))
    if len(buttons) != 1:
        raise LookupError(f"ECM Search options control must be visible exactly once; observed {len(buttons)}.")
    button = buttons[0]
    button_id = await button.get_attribute("id")
    if not button_id or not button_id.endswith("_moreOptionsLink"):
        raise LookupError("ECM Search options dialog cannot be identified safely.")
    page = button.page
    dialog_id = button_id.removesuffix("_moreOptionsLink") + "_moreOptionsDialog"
    dialogs = page.locator(f'[role="dialog"][id="{dialog_id}"]')
    visible_dialogs = await _visible_locators(dialogs)
    if not visible_dialogs:
        await click_control(button, timeout_ms)
    deadline = asyncio.get_running_loop().time() + timeout_ms / 1000
    while not visible_dialogs and asyncio.get_running_loop().time() < deadline:
        visible_dialogs = await _visible_locators(dialogs)
        if not visible_dialogs:
            await asyncio.sleep(0.1)
    if len(visible_dialogs) != 1:
        raise LookupError(f"ECM Search options dialog must be visible exactly once; observed {len(visible_dialogs)}.")
    dialog = visible_dialogs[0]
    rows = []
    while not rows and asyncio.get_running_loop().time() < deadline:
        rows = await _visible_locators(dialog.locator('tr[data-dojo-attach-point="_versionOptionNode"]'))
        if not rows:
            await asyncio.sleep(0.1)
    if len(rows) != 1:
        raise LookupError(f"ECM Version selector must be visible exactly once; observed {len(rows)}.")
    row = rows[0]
    listboxes = []
    while not listboxes and asyncio.get_running_loop().time() < deadline:
        listboxes = await _visible_locators(row.get_by_role("listbox"))
        if not listboxes:
            await asyncio.sleep(0.1)
    if len(listboxes) != 1:
        raise LookupError(f"ECM Version selector control must be visible exactly once; observed {len(listboxes)}.")

    async def selected_state() -> tuple[str, str]:
        try:
            state = await listboxes[0].evaluate("""element => ({
                label: element.querySelector('[role=option][aria-selected=true]')?.innerText?.trim() || '',
                value: element.querySelector('input[type=hidden][data-dojo-attach-point=valueNode]')?.value || ''
            })""")
            return str(state.get("label", "")).strip(), str(state.get("value", "")).strip()
        except Exception as error:
            raise LookupError("ECM Version selector state could not be read safely.") from error

    label, value = await selected_state()
    if label.casefold() != "all versions" or value.casefold() != "allversions":
        await click_control(listboxes[0], timeout_ms)
        options = []
        deadline = asyncio.get_running_loop().time() + timeout_ms / 1000
        while not options and asyncio.get_running_loop().time() < deadline:
            options = await _visible_locators(page.get_by_role("option", name="All versions", exact=True))
            if not options:
                await asyncio.sleep(0.1)
        if len(options) != 1:
            raise LookupError(f"Visible ECM option 'All versions' must appear exactly once; observed {len(options)}.")
        await click_control(options[0], timeout_ms)
        deadline = asyncio.get_running_loop().time() + timeout_ms / 1000
        while asyncio.get_running_loop().time() < deadline:
            label, value = await selected_state()
            if label.casefold() == "all versions" and value.casefold() == "allversions":
                break
            await asyncio.sleep(0.1)
    if label.casefold() != "all versions" or value.casefold() != "allversions":
        raise LookupError(f"ECM Version must be 'All versions' (allversions); observed {label!r} ({value!r}).")
    ok_buttons = await _visible_locators(dialog.get_by_role("button", name="OK", exact=True))
    if len(ok_buttons) != 1:
        raise LookupError(f"ECM Search options OK button must be visible exactly once; observed {len(ok_buttons)}.")
    await click_control(ok_buttons[0], timeout_ms)
    deadline = asyncio.get_running_loop().time() + timeout_ms / 1000
    while asyncio.get_running_loop().time() < deadline:
        if not await _visible_locators(dialogs) and "all versions" in (await button.inner_text()).casefold():
            break
        await asyncio.sleep(0.1)
    else:
        raise LookupError("ECM Search options did not apply All versions; refusing to search.")
    _log("ecm_version", label=label, value=value)


async def _lookup_results_snapshot(panel, page) -> dict[str, object]:
    """Capture the visible results pane and busy flag for per-search freshness checks."""
    state = await page.evaluate("""() => {
        const visible = element => !!element && !!(element.offsetWidth || element.offsetHeight || element.getClientRects().length);
        const panes = Array.from(document.querySelectorAll('.searchResultsPane')).filter(visible);
        const pane = panes[panes.length - 1];
        if (!pane) return null;
        if (window.__v3LookupResultsObserver?.pane !== pane) {
            window.__v3LookupResultsObserver?.observer?.disconnect();
            const tracker = {pane, mutations: 0, observer: null};
            tracker.observer = new MutationObserver(records => {
                if (records.length) tracker.mutations += records.length;
            });
            tracker.observer.observe(pane, {subtree: true, childList: true, characterData: true, attributes: true});
            window.__v3LookupResultsObserver = tracker;
        }
        const emptyNode = pane.querySelector('[data-dojo-attach-point="emptySearchResultsPane"]');
        const emptyText = visible(emptyNode) ? (emptyNode.innerText || emptyNode.textContent || '').trim() : '';
        const rows = Array.from(pane.querySelectorAll('[role="row"]')).filter(visible)
            .map(row => (row.innerText || row.textContent || '').trim());
        return {
            busy: !!(document.body && document.body.classList.contains('ecmWait')),
            mutations: window.__v3LookupResultsObserver.mutations,
            text: (pane.innerText || pane.textContent || '').replace(/\\s+/g, ' ').trim(),
            rows,
            emptyText
        };
    }""")
    if not isinstance(state, dict):
        raise LookupError("ECM Search Results pane is not available for a freshness check.")
    return state


async def _wait_for_lookup_results(panel, page, timeout_ms: int, before: dict[str, object]) -> bool:
    """Accept only a completed post-click update, never an old visible grid."""
    deadline = asyncio.get_running_loop().time() + timeout_ms / 1000
    saw_busy = False
    before_signature = (
        before.get("mutations"), before.get("text"), tuple(before.get("rows", ())), before.get("emptyText")
    )
    while asyncio.get_running_loop().time() < deadline:
        current = await _lookup_results_snapshot(panel, page)
        saw_busy = saw_busy or bool(current.get("busy"))
        signature = (
            current.get("mutations"), current.get("text"), tuple(current.get("rows", ())), current.get("emptyText")
        )
        changed = signature != before_signature
        idle = not bool(current.get("busy"))
        empty_text = str(current.get("emptyText", "")).strip()
        initial_empty = not empty_text or "run a search to see results" in empty_text.casefold()
        result_rows = list(current.get("rows", ()))
        has_rows = len(result_rows) > 1
        completed_empty = not initial_empty and bool(empty_text)
        if idle and (saw_busy or changed) and (has_rows or completed_empty):
            return not completed_empty
        await asyncio.sleep(0.2)
    raise LookupError("ECM search could not be verified as a fresh completed result; refusing to use a visible old grid.")


async def _read_oldest_result_date(panel, page, timeout_ms: int, *, before: dict[str, object], query_id: int, row_number: int, outgoing_reference: str) -> date | None:
    if not await _wait_for_lookup_results(panel, page, timeout_ms, before):
        _log("lookup_results", query_id=query_id, row=row_number, outgoing_reference=outgoing_reference, result_count=0, dates=[])
        return None
    toolbar = panel.get_by_role("toolbar", name="SearchResultsToolbar")
    if callable(getattr(panel, "locator", None)) and not await _visible_locators(toolbar):
        toolbar = panel.get_by_role("toolbar")
    body = await page.locator("body").inner_text()
    count_matches = re.findall(r"The search returned\s+([\d,]+)\s+items?", body, re.IGNORECASE)
    if count_matches and int(count_matches[-1].replace(",", "")) == 0:
        _log("lookup_results", query_id=query_id, row=row_number, outgoing_reference=outgoing_reference, result_count=0, dates=[])
        return None
    grid = panel.get_by_role("grid", name="grid")
    date_index = None
    headers = await _visible_locators(grid.get_by_role("columnheader"))
    for index, header in enumerate(headers):
        header_text = re.sub(r"^[\s▲▼↑↓]+|[\s▲▼↑↓]+$", "", await header.inner_text())
        if header_text.casefold() == "document date":
            date_index = index
            break
    if date_index is None:
        raise LookupError("ECM results do not expose a Document date column.")
    dates: list[date] = []
    result_ordinal = 0
    seen_pages: set[tuple[str, ...]] = set()
    while True:
        rows = await grid.get_by_role("row").all()
        page_signature = tuple([(await row.inner_text()).strip() for row in rows[1:]])
        if page_signature in seen_pages:
            raise LookupError("ECM result pagination repeated; refusing an incomplete oldest-date result.")
        seen_pages.add(page_signature)
        for row in rows[1:]:
            cells = await row.get_by_role("gridcell").all()
            if len(cells) <= date_index:
                continue
            value = (await cells[date_index].inner_text()).strip()
            if value:
                result_date = _date_from_result(value)
                dates.append(result_date)
                result_ordinal += 1
                _log("lookup_result_date", query_id=query_id, row=row_number, outgoing_reference=outgoing_reference, ordinal=result_ordinal, date=result_date.isoformat())
        next_controls = await _visible_locators(toolbar.get_by_role("button", name=re.compile("next|more", re.IGNORECASE)))
        enabled = []
        for button in next_controls:
            if await button.is_enabled() and (await button.get_attribute("aria-disabled")) != "true":
                enabled.append(button)
        if len(enabled) > 1:
            raise LookupError("ECM result pagination has ambiguous next controls.")
        if not enabled:
            break
        previous_signature = page_signature
        before_page = await _lookup_results_snapshot(panel, page)
        await enabled[0].click(timeout=timeout_ms)
        page_deadline = asyncio.get_running_loop().time() + timeout_ms / 1000
        while asyncio.get_running_loop().time() < page_deadline:
            await asyncio.sleep(0.1)
            updated_rows = await grid.get_by_role("row").all()
            updated_signature = tuple([(await row.inner_text()).strip() for row in updated_rows[1:]])
            after_page = await _lookup_results_snapshot(panel, page)
            if (
                updated_signature != previous_signature
                or after_page.get("mutations", 0) > before_page.get("mutations", 0)
            ):
                break
        else:
            raise LookupError("ECM result pagination did not advance to a fresh page.")
        _log("lookup_results_page", query_id=query_id, row=row_number, page=len(seen_pages) + 1)
    result_count = int(count_matches[-1].replace(",", "")) if count_matches else result_ordinal
    _log("lookup_results", query_id=query_id, row=row_number, outgoing_reference=outgoing_reference, result_count=result_count, dates=[item.isoformat() for item in dates])
    return min(dates) if dates else None


async def run_lookup(*, workbook_path: Path, worksheet_name: str, csv_path: Path, profile_dir: Path, ecm_url: str, timeout_seconds: int, summary_path: Path | None = None, diagnostics_dir: Path | None = None, reset_profile_requested: bool = False, scope: str = "all", previous_summary_path: Path | None = None) -> tuple[int, int, int]:
    paths = [(workbook_path, "Workbook"), (csv_path, "CSV"), (profile_dir, "Edge profile")]
    if summary_path is not None:
        paths.append((summary_path, "summary"))
    if previous_summary_path is not None:
        paths.append((previous_summary_path, "previous summary"))
    for path, label in paths:
        try:
            path.resolve().relative_to(Path(__file__).resolve().parents[1])
        except ValueError as error:
            raise LookupError(f"V3 {label} path must remain under control_panel_v3: {path}") from error
    if not workbook_path.is_file() or not csv_path.is_file():
        raise LookupError("V3 review workbook and Step 0 CSV must both exist.")
    if not _inside(profile_dir, Path(__file__).resolve().parents[1]):
        raise LookupError(f"V3 Edge profile path must remain under control_panel_v3: {profile_dir}")
    if workbook_path.with_name(f"~${workbook_path.name}").exists():
        raise LookupError("Close the V3 review workbook in Excel before running the lookup.")
    if diagnostics_dir is not None and not _inside(diagnostics_dir, Path(__file__).resolve().parents[1]):
        raise LookupError(f"V3 diagnostics path must remain under control_panel_v3: {diagnostics_dir}")
    if reset_profile_requested:
        backup = reset_profile(profile_dir)
        _log("profile_reset", backup_profile=backup)
    reference_map = read_outgoing_reference_map(csv_path)
    workbook = load_workbook(workbook_path)
    try:
        if worksheet_name not in workbook.sheetnames:
            raise LookupError(f"Review worksheet was not found: {worksheet_name}")
        worksheet = workbook[worksheet_name]
        validate_workbook(worksheet)
        prior_rows: set[int] | None = None
        if scope == "previously-unmatched":
            if previous_summary_path is None or not previous_summary_path.is_file():
                raise LookupError("Previously-unmatched scope requires an existing prior lookup summary.")
            try:
                previous = json.loads(previous_summary_path.read_text(encoding="utf-8"))
                prior_rows = {int(item["row"]) for item in previous.get("misses", [])}
            except (OSError, json.JSONDecodeError, TypeError, KeyError, ValueError) as error:
                raise LookupError(f"Prior lookup summary is invalid: {error}") from error
        tasks = build_search_rows(worksheet, reference_map, scope=scope, previously_unmatched=prior_rows)
        _log("lookup_started", rows=len(tasks), worksheet=worksheet_name, scope=scope)
        if not tasks:
            _log("lookup_complete", hits=0, misses=0, searched=0, scope=scope)
            _write_summary(summary_path, worksheet=worksheet_name, hits=0, misses=[], searched=0, scope=scope)
            return 0, 0, 0
        try:
            from playwright.async_api import async_playwright
        except ImportError as error:
            raise LookupError("Playwright is required in the V3 environment.") from error
        hits = 0
        misses: list[dict[str, object]] = []
        browser_diagnostics = BrowserDiagnostics()
        page = None
        async with async_playwright() as playwright:
            context = await playwright.chromium.launch_persistent_context(user_data_dir=str(profile_dir), channel="msedge", headless=False, accept_downloads=False)
            try:
                page = context.pages[0] if context.pages else await context.new_page()
                browser_diagnostics.attach(page)
                await page.goto(ecm_url, wait_until="domcontentloaded", timeout=timeout_seconds * 1000)
                panel = await _open_lookup_with_auth(page, timeout_seconds)
                for index, (row_number, reference) in enumerate(tasks):
                    try:
                        if index > 0:
                            # Re-verify and expand the Search Criteria section afresh before every subsequent reference.
                            panel = await _open_lookup_with_auth(page, timeout_seconds)
                        await _ensure_lookup_criteria_expanded(panel, timeout_seconds * 1000)
                        await _ensure_all_versions(panel, timeout_seconds * 1000)
                        form = await _active_form(panel)
                        await _set_search_criteria(form, reference)
                        search_buttons = await _visible_locators(panel.get_by_role("button", name="Search", exact=True))
                        if len(search_buttons) != 1:
                            raise LookupError("ECM Search button was not found uniquely.")
                        before = await _lookup_results_snapshot(panel, page)
                        query_id = index + 1
                        _log("lookup_query", query_id=query_id, row=row_number, outgoing_reference=reference, version="All versions")
                        await click_control(search_buttons[0], timeout_seconds * 1000)
                        first_date = await _read_oldest_result_date(
                            panel, page, timeout_seconds * 1000, before=before,
                            query_id=query_id, row_number=row_number, outgoing_reference=reference,
                        )
                    except Exception as error:
                        raise LookupError(f"ECM lookup failed for workbook row {row_number}: {error}") from error
                    if first_date is None:
                        misses.append({"row": row_number, "outgoing_reference": reference})
                        _log("lookup_miss", row=row_number, outgoing_reference=reference)
                        continue
                    worksheet.cell(row_number, REMARK_COLUMN).value = first_date.isoformat()
                    _atomic_save(workbook, workbook_path, row_number)
                    hits += 1
                    _log("lookup_hit", row=row_number, outgoing_reference=reference, first_date=first_date.isoformat())
            except Exception:
                if diagnostics_dir is not None and page is not None:
                    try:
                        await save_failure_diagnostics(page, diagnostics_dir, "lookup", browser_diagnostics)
                    except Exception as diagnostic_error:
                        _log("diagnostics_capture_failed", error=type(diagnostic_error).__name__)
                raise
            finally:
                await context.close()
        _log("lookup_complete", hits=hits, misses=len(misses), searched=len(tasks))
        _write_summary(summary_path, worksheet=worksheet_name, hits=hits, misses=misses, searched=len(tasks), scope=scope)
        return hits, len(misses), len(tasks)
    finally:
        workbook.close()


def _parse_date(value: str) -> date:
    return date.fromisoformat(value)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Find first ECM document dates for V3 workbook rows.")
    parser.add_argument("--workbook", required=True, type=Path)
    parser.add_argument("--worksheet", required=True)
    parser.add_argument("--csv", required=True, type=Path)
    parser.add_argument("--profile-dir", required=True, type=Path)
    parser.add_argument("--ecm-url", default=os.getenv("CONTROL_PANEL_V3_ECM_URL", DEFAULT_ECM_URL))
    parser.add_argument("--timeout-seconds", type=int, default=120)
    parser.add_argument("--summary", type=Path)
    parser.add_argument("--diagnostics-dir", type=Path)
    parser.add_argument("--reset-profile", action="store_true")
    parser.add_argument("--scope", choices=("all", "blank", "previously-unmatched"), default="all")
    parser.add_argument("--previous-summary", type=Path)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.timeout_seconds <= 0:
        raise LookupError("Timeout must be greater than zero.")
    hits, misses, rows = asyncio.run(run_lookup(workbook_path=args.workbook, worksheet_name=args.worksheet, csv_path=args.csv, profile_dir=args.profile_dir, ecm_url=args.ecm_url, timeout_seconds=args.timeout_seconds, summary_path=args.summary, diagnostics_dir=args.diagnostics_dir, reset_profile_requested=args.reset_profile, scope=args.scope, previous_summary_path=args.previous_summary))
    print(f"Lookup complete: {hits} matched, {misses} unmatched, {rows} searched.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (LookupError, OSError, ValueError) as error:
        print(f"Error: {error}", file=sys.stderr)
        raise SystemExit(1)
