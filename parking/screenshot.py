"""Headless screenshots of the booking calendar, for Telegram proof-of-state.

Reuses ``login.py``'s persistent, already-authenticated browser profile and
stealth launch conventions - no new login step, no separate token handling.
Capture is best-effort throughout: any failure just means the caller falls
back to a text-only notification instead of losing the alert.

The site (a single-page app) has no URL params for center or month - both are
client-only UI state, confirmed by driving the real page. The ``Center``
selector (``#center``) opens a listbox of the configured center names; picking
one resets the calendar to the current month. An accessible ``"Next Month"``
button advances it. Never click anything else on this form - the same page
also carries the real booking ``Save`` button.
"""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import List

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

from parking.login import STEALTH_INIT_SCRIPT, stealth_launch_kwargs

logger = logging.getLogger("parking.screenshot")

WAIT_AFTER_LOAD_MS = 3000
WAIT_AFTER_CALENDAR_UPDATE_MS = 1200


def _slug(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "-", name).strip("-").lower() or "site"


def capture_all(centers: List[str], site_url: str, out_dir: Path, profile_dir: Path) -> List[Path]:
    """Screenshot each center's calendar, this month and next.

    Returns the PNGs actually captured, in ``[center1-this, center1-next,
    center2-this, center2-next, ...]`` order. Best-effort per center: a
    failure switching to or screenshotting one center is logged and skipped,
    so the others are still captured. Never raises - a locked profile, a
    navigation failure, or any Playwright error returns whatever was captured
    so far (possibly nothing).
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    captured: List[Path] = []
    try:
        with sync_playwright() as pw:
            context = pw.chromium.launch_persistent_context(
                **stealth_launch_kwargs(str(profile_dir), headless=True))
            try:
                context.add_init_script(STEALTH_INIT_SCRIPT)
                page = context.pages[0] if context.pages else context.new_page()
                page.goto(site_url, wait_until="networkidle")
                page.wait_for_timeout(WAIT_AFTER_LOAD_MS)
                for center in centers:
                    try:
                        page.click("#center")
                        page.get_by_role("option", name=center, exact=True).click()
                        page.wait_for_timeout(WAIT_AFTER_CALENDAR_UPDATE_MS)
                        this_month = out_dir / f"{_slug(center)}-this-month.png"
                        page.screenshot(path=str(this_month))
                        captured.append(this_month)

                        page.get_by_role("button", name="Next Month").click()
                        page.wait_for_timeout(WAIT_AFTER_CALENDAR_UPDATE_MS)
                        next_month = out_dir / f"{_slug(center)}-next-month.png"
                        page.screenshot(path=str(next_month))
                        captured.append(next_month)
                    except PlaywrightError as exc:
                        logger.warning("⚠️ screenshot capture failed for center %s: %s", center, exc)
                        continue
            finally:
                context.close()
    except (PlaywrightError, OSError) as exc:
        logger.warning("⚠️ screenshot capture failed: %s", exc)
    return captured
