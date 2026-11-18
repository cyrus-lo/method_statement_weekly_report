"""V3-only browser pacing, profile reset, and failure diagnostics helpers."""

from __future__ import annotations

import asyncio
import json
import random
import re
import shutil
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse


class PageCursor:
    """Move only Playwright's page-scoped pointer before a locator interaction."""

    def __init__(self, page, *, rng: random.Random | None = None) -> None:
        self.page = page
        self.rng = rng or random.Random()
        self.position = (0.0, 0.0)

    async def move_to(self, locator) -> bool:
        try:
            box = await locator.bounding_box()
            if not box or box.get("width", 0) <= 0 or box.get("height", 0) <= 0:
                return False
            target = (
                box["x"] + box["width"] * self.rng.uniform(0.35, 0.65),
                box["y"] + box["height"] * self.rng.uniform(0.35, 0.65),
            )
            distance = ((target[0] - self.position[0]) ** 2 + (target[1] - self.position[1]) ** 2) ** 0.5
            steps = max(8, min(24, round(distance / 35)))
            duration = min(0.70, 0.18 + 0.52 * min(1.0, distance / 900))
            pause_indices = set(self.rng.sample(range(max(1, steps - 1)), self.rng.randint(0, min(3, steps - 1))))
            for index in range(1, steps + 1):
                point = (self.position[0] + (target[0] - self.position[0]) * index / steps,
                         self.position[1] + (target[1] - self.position[1]) * index / steps)
                await self.page.mouse.move(*point)
                await asyncio.sleep(duration / steps)
                if index - 1 in pause_indices:
                    pause = self.rng.uniform(0.40, 1.50) if self.rng.random() < 0.16 else self.rng.uniform(0.04, 0.18)
                    await asyncio.sleep(min(1.50, pause))
            self.position = target
            return True
        except Exception:
            return False


async def click_control(locator, timeout_ms: int) -> None:
    """Prefer accessible locator activation; use its DOM click event as fallback."""
    page = getattr(locator, "page", None)
    if page is not None:
        cursor = getattr(page, "_v3_page_cursor", None)
        if cursor is None:
            cursor = PageCursor(page)
            try:
                setattr(page, "_v3_page_cursor", cursor)
            except Exception:
                pass
        await cursor.move_to(locator)
    try:
        await locator.click(timeout=min(timeout_ms, 10_000))
    except Exception:
        try:
            await locator.dispatch_event("click")
        except Exception as error:
            raise RuntimeError(f"ECM control activation failed ({type(error).__name__}).") from error


def reset_profile(profile_dir: Path) -> Path | None:
    """Back up a persistent V3 profile before an explicitly requested reset."""
    if not profile_dir.exists():
        return None
    stamp = datetime.now().astimezone().strftime("%Y%m%d-%H%M%S")
    backup = profile_dir.with_name(f"{profile_dir.name}.backup-{stamp}")
    suffix = 1
    while backup.exists():
        backup = profile_dir.with_name(f"{profile_dir.name}.backup-{stamp}-{suffix}")
        suffix += 1
    try:
        profile_dir.rename(backup)
    except OSError as error:
        raise RuntimeError(f"Could not back up the V3 Edge profile; close Edge using it and retry: {profile_dir}") from error
    return backup


class BrowserDiagnostics:
    """Bounded browser failure metadata that excludes page text and URL queries."""

    def __init__(self, limit: int = 100) -> None:
        self.limit = limit
        self.events: list[dict[str, object]] = []

    @staticmethod
    def _safe_url(value: object) -> str:
        try:
            parsed = urlparse(str(value))
            return parsed._replace(query="", fragment="").geturl()
        except Exception:
            return "<unavailable>"

    def attach(self, page) -> None:
        if not callable(getattr(page, "on", None)):
            return

        def record(kind: str, **fields: object) -> None:
            self.events.append({"kind": kind, **fields})
            del self.events[:-self.limit]

        def on_console(message) -> None:
            level = str(getattr(message, "type", ""))
            if level in {"warning", "error"}:
                record("console", level=level)

        def on_page_error(error) -> None:
            record("page_error", error_type=type(error).__name__)

        def on_request_failed(request) -> None:
            record("request_failed", method=str(getattr(request, "method", "")), resource_type=str(getattr(request, "resource_type", "")), url=self._safe_url(getattr(request, "url", "")))

        def on_response(response) -> None:
            status = int(getattr(response, "status", 0) or 0)
            if status >= 400:
                request = getattr(response, "request", None)
                record("response_error", status=status, method=str(getattr(request, "method", "")), url=self._safe_url(getattr(response, "url", "")))

        page.on("console", on_console)
        page.on("pageerror", on_page_error)
        page.on("requestfailed", on_request_failed)
        page.on("response", on_response)


async def save_failure_diagnostics(page, directory: Path, stage: str, diagnostics: BrowserDiagnostics) -> list[Path]:
    """Capture sensitive DOM/screenshot plus safe bounded metadata for a failed run."""
    directory.mkdir(parents=True, exist_ok=True)
    label = re.sub(r"[^a-zA-Z0-9_-]+", "_", stage).strip("_") or "unknown"
    screenshot = directory / f"failure-{label}.png"
    dom = directory / f"failure-{label}.html"
    metadata = directory / f"failure-{label}.diagnostics.json"
    paths = []
    try:
        await page.screenshot(path=str(screenshot), full_page=True)
        paths.append(screenshot)
    except Exception:
        pass
    try:
        dom.write_text(await page.content(), encoding="utf-8")
        paths.append(dom)
    except Exception:
        pass
    try:
        title = await page.title()
    except Exception:
        title = "<unavailable>"
    try:
        parsed = urlparse(str(getattr(page, "url", "")))
        safe_url = parsed._replace(query="", fragment="").geturl()
    except Exception:
        safe_url = "<unavailable>"
    metadata.write_text(json.dumps({"stage": stage, "url": safe_url, "title": title, "browser_events": diagnostics.events}, ensure_ascii=True, indent=2), encoding="utf-8")
    paths.append(metadata)
    return paths


async def wait_for_manual_auth(page, is_authenticated, ready, timeout_seconds: int) -> None:
    deadline = asyncio.get_running_loop().time() + timeout_seconds
    while asyncio.get_running_loop().time() < deadline:
        if await is_authenticated(page):
            try:
                await ready(page, 2_000)
                return
            except Exception:
                pass
        await asyncio.sleep(1)
    raise RuntimeError("Manual ECM SSO/MFA authentication timed out; complete login in the dedicated V3 Edge window and retry.")
