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
from .ecm_step0_export import DEFAULT_ECM_URL, _active_form, _control_metadata, _control_value, _labeled_controls, _open_saved_search, _page_requires_authentication, _visible_locators
from .properties_csv import CSV_HEADERS, parse_ecm_datetime

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
        reference = reference_map[source_identity]
        remark = str(worksheet.cell(row_number, REMARK_COLUMN).value or "").strip()
        if scope == "blank" and remark:
            continue
        if scope == "previously-unmatched" and (row_number not in (previously_unmatched or set()) or remark):
            continue
        if reference:
            tasks.append((row_number, reference))
        else:
            _log("lookup_miss", row=row_number, reason="blank_outgoing_reference")
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


async def _set_search_criteria(form, outgoing_reference: str) -> None:
    controls = await _labeled_controls(form, "In/Out")
    matches = [control for control in controls if "out" == (await _control_value(control)).casefold().strip()]
    if len(matches) != 1:
        raise LookupError(f"ECM criterion 'In/Out' must be exactly 'OUT'; observed {len(matches)} matching controls.")
    ref_controls = await _labeled_controls(form, "Outgoing ref. no.", include_operator=True)
    operator_controls = []
    value_controls = []
    for control in ref_controls:
        metadata = await _control_metadata(control)
        if "Operator" in metadata.get("id", "") or "operatorsSelect" in metadata.get("class_name", ""):
            operator_controls.append(control)
        else:
            value_controls.append(control)
    operator_matches = [control for control in operator_controls if (await _control_value(control)).casefold().strip() == "starts with"]
    if len(operator_matches) != 1:
        raise LookupError("ECM Outgoing ref. no. operator must be exactly 'Starts With'.")
    if len(value_controls) != 1:
        raise LookupError(f"ECM Outgoing ref. no. value control is missing or ambiguous ({len(value_controls)} found).")
    await value_controls[0].fill(outgoing_reference)
    await value_controls[0].dispatch_event("change")
    await value_controls[0].press("Tab")


async def _read_oldest_result_date(panel, page, timeout_ms: int) -> date | None:
    toolbar = panel.get_by_role("toolbar", name="SearchResultsToolbar")
    await toolbar.first.wait_for(state="visible", timeout=timeout_ms)
    body = await page.locator("body").inner_text()
    count_match = re.search(r"The search returned\s+([\d,]+)\s+items?", body, re.IGNORECASE)
    if count_match and int(count_match.group(1).replace(",", "")) == 0:
        return None
    grid = panel.get_by_role("grid", name="grid")
    date_index = None
    headers = await _visible_locators(grid.get_by_role("columnheader"))
    for index, header in enumerate(headers):
        if (await header.inner_text()).strip().casefold() == "document date":
            date_index = index
            break
    if date_index is None:
        raise LookupError("ECM results do not expose a Document date column.")
    dates: list[date] = []
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
                dates.append(_date_from_result(value))
        next_controls = await _visible_locators(toolbar.get_by_role("button", name=re.compile("next|more", re.IGNORECASE)))
        enabled = []
        for button in next_controls:
            if await button.is_enabled() and (await button.get_attribute("aria-disabled")) != "true":
                enabled.append(button)
        if len(enabled) > 1:
            raise LookupError("ECM result pagination has ambiguous next controls.")
        if not enabled:
            break
        await enabled[0].click(timeout=timeout_ms)
        await asyncio.sleep(0.25)
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
                if await _page_requires_authentication(page):
                    await wait_for_manual_auth(
                        page,
                        _page_requires_authentication,
                        lambda current_page, timeout: _open_saved_search(current_page, timeout, search_name=LOOKUP_SEARCH_NAME),
                        timeout_seconds,
                    )
                panel = await _open_saved_search(page, timeout_seconds * 1000, search_name=LOOKUP_SEARCH_NAME)
                for row_number, reference in tasks:
                    try:
                        form = await _active_form(panel)
                        await _set_search_criteria(form, reference)
                        search_buttons = await _visible_locators(panel.get_by_role("button", name="Search", exact=True))
                        if len(search_buttons) != 1:
                            raise LookupError("ECM Search button was not found uniquely.")
                        await click_control(search_buttons[0], timeout_seconds * 1000)
                        first_date = await _read_oldest_result_date(panel, page, timeout_seconds * 1000)
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
                    # Use the application's Search page controls afresh for every reference.
                    panel = await _open_saved_search(page, timeout_seconds * 1000, search_name=LOOKUP_SEARCH_NAME)
            finally:
                await context.close()
        _log("lookup_complete", hits=hits, misses=len(misses), searched=len(tasks))
        _write_summary(summary_path, worksheet=worksheet_name, hits=hits, misses=misses, searched=len(tasks), scope=scope)
        return hits, len(misses), len(tasks)
    except Exception:
        if diagnostics_dir is not None and locals().get("page") is not None:
            await save_failure_diagnostics(locals()["page"], diagnostics_dir, "lookup", browser_diagnostics)
        raise
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
