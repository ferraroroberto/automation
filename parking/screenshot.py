"""Headless screenshot of the parking site, for Telegram proof-of-state.

Reuses ``login.py``'s persistent, already-authenticated browser profile and
stealth launch conventions - no new login step, no separate token handling.
Capture is best-effort: any failure just means the caller falls back to a
text-only notification instead of losing the alert.
"""

from __future__ import annotations

import logging
from pathlib import Path

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

from parking.login import STEALTH_INIT_SCRIPT, stealth_launch_kwargs

logger = logging.getLogger("parking.screenshot")

WAIT_AFTER_LOAD_MS = 3000


def capture(site_url: str, out_path: Path, profile_dir: Path) -> bool:
    """Screenshot ``site_url`` (via the logged-in profile) to ``out_path``.

    Headless, using the same persistent profile ``login.py`` authenticates.
    Never raises: a locked profile, a navigation failure, or any Playwright
    error is logged and reported as False.
    """
    try:
        with sync_playwright() as pw:
            context = pw.chromium.launch_persistent_context(
                **stealth_launch_kwargs(str(profile_dir), headless=True))
            try:
                context.add_init_script(STEALTH_INIT_SCRIPT)
                page = context.pages[0] if context.pages else context.new_page()
                page.goto(site_url, wait_until="networkidle")
                page.wait_for_timeout(WAIT_AFTER_LOAD_MS)
                out_path.parent.mkdir(parents=True, exist_ok=True)
                page.screenshot(path=str(out_path), full_page=True)
                return True
            finally:
                context.close()
    except (PlaywrightError, OSError) as exc:
        logger.warning("⚠️ screenshot capture failed: %s", exc)
        return False
