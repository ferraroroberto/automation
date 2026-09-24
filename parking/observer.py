"""Headed Chrome windows on the bookings list, so the owner can watch a burst (#139).

Watch-only: nothing here clicks. Each window loads before the target and is
reloaded once, when its day's worker finishes - never during the fast loop, so
the windows add no load on the site at the release moment. One Chrome on the
persistent login profile (a profile can't be opened by two Chrome instances),
one window per day, cascaded so each stays visible.

Best-effort throughout: any Playwright failure is logged and the burst carries on
without windows - they are only a view, the workers book through the API.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Dict, List, Optional

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

from parking.login import STEALTH_INIT_SCRIPT, stealth_launch_kwargs

logger = logging.getLogger("parking.observer")

CASCADE_PX = 70  # offset between windows, so every title bar stays visible


def bookings_url(site_url: str) -> str:
    """The bookings list: `PARKING_URL` is the booking form (`.../Booking/create`)."""
    return site_url[: -len("/create")] if site_url.endswith("/create") else site_url


class Observers:
    def __init__(self, playwright, context, pages: Dict[str, object]) -> None:
        self._playwright = playwright
        self._context = context
        self._pages = pages

    @classmethod
    def open(cls, site_url: str, days: List[str], profile_dir: Path) -> Optional["Observers"]:
        """One headed window per day on the bookings list, or None if Chrome can't be opened."""
        if not site_url or not days:
            return None
        url = bookings_url(site_url)
        playwright = sync_playwright().start()
        try:
            context = playwright.chromium.launch_persistent_context(
                **stealth_launch_kwargs(str(profile_dir), headless=False))
            context.add_init_script(STEALTH_INIT_SCRIPT)
            first = context.pages[0] if context.pages else context.new_page()
            cdp = context.new_cdp_session(first)
            pages: Dict[str, object] = {days[0]: first}
            for day in days[1:]:
                # new_page() would only add a tab; a CDP target with newWindow is its own window.
                with context.expect_page() as opened:
                    cdp.send("Target.createTarget", {"url": "about:blank", "newWindow": True})
                pages[day] = opened.value
            for index, (day, page) in enumerate(pages.items()):
                _place(context, page, index)
                page.goto(url, wait_until="domcontentloaded")
                _title(page, day)
            logger.info("ℹ️ observer windows open for %s", ", ".join(days))
            return cls(playwright, context, pages)
        except (PlaywrightError, OSError) as exc:
            logger.warning("⚠️ observer windows unavailable (%s); the burst goes on without them", exc)
            playwright.stop()
            return None

    def refresh(self, day: str) -> None:
        page = self._pages.get(day)
        if page is None:
            return
        try:
            page.reload(wait_until="domcontentloaded")
            _title(page, day)
        except PlaywrightError as exc:
            logger.warning("⚠️ could not refresh the %s window: %s", day, exc)

    def close(self) -> None:
        try:
            self._context.close()
        except PlaywrightError as exc:
            logger.warning("⚠️ closing the observer windows failed: %s", exc)
        finally:
            self._playwright.stop()


def _place(context, page, index: int) -> None:
    session = context.new_cdp_session(page)
    window = session.send("Browser.getWindowForTarget")["windowId"]
    offset = CASCADE_PX * index
    session.send("Browser.setWindowBounds",
                 {"windowId": window, "bounds": {"left": offset, "top": offset, "windowState": "normal"}})


def _title(page, day: str) -> None:
    """Name each window after its day, so three identical lists can be told apart."""
    try:
        page.evaluate("day => { document.title = `Parking burst - ${day}`; }", day)
    except PlaywrightError:
        pass  # a title is cosmetic; the page may still be navigating
