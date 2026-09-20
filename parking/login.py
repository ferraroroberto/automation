"""One-time (about monthly) manual login that saves the site token to .env.

``python -m parking.login``   opens a dedicated Chrome profile; you sign in by hand
                              (SSO + any 2FA); the token is read from the page's
                              localStorage and written to .env as PARKING_TOKEN.
``python -m parking.login --refresh``  headless: re-reads the token from the saved profile.

The token is never printed or logged; only its expiry is.
"""

from __future__ import annotations

import argparse
import logging
import sys
import time
from datetime import datetime, timezone
from typing import Optional

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

from parking import auth
from parking.config import PROFILE_DIR, load_env, set_env_value

logger = logging.getLogger("parking.login")

# Same real-Chrome, human-looking launch conventions as the other fleet browser automations.
STEALTH_INIT_SCRIPT = "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"


def stealth_launch_kwargs(user_data_dir: str, headless: bool) -> dict:
    return {
        "user_data_dir": user_data_dir,
        "channel": "chrome",
        "headless": headless,
        "locale": "en-US",
        "args": [
            "--disable-blink-features=AutomationControlled",
            "--disable-features=Translate",
            "--no-default-browser-check",
            "--no-first-run",
            "--lang=en-US",
        ],
        "ignore_default_args": ["--enable-automation", "--enable-blink-features=IdleDetection"],
        "viewport": {"width": 1280, "height": 900},
        "chromium_sandbox": True,
    }


def _read_token(context) -> Optional[str]:
    """Token from any open page; pages mid-navigation (SSO redirects) or on other origins are skipped."""
    for page in context.pages:
        try:
            token = auth.normalize_token(page.evaluate("() => window.localStorage.getItem('token')"))
        except PlaywrightError:
            continue
        if token:
            return token
    return None


def grab_token(site_url: str, headless: bool, wait_seconds: int) -> Optional[str]:
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as pw:
        context = pw.chromium.launch_persistent_context(
            **stealth_launch_kwargs(str(PROFILE_DIR), headless))
        try:
            context.add_init_script(STEALTH_INIT_SCRIPT)
            page = context.pages[0] if context.pages else context.new_page()
            page.goto(site_url)
            deadline = time.monotonic() + wait_seconds
            while time.monotonic() < deadline:
                token = _read_token(context)
                expiry = auth.jwt_expiry(token) if token else None
                if token and expiry and expiry > datetime.now(timezone.utc):
                    return token
                time.sleep(2)
            return None
        finally:
            context.close()


def main(argv: Optional[list] = None) -> int:
    parser = argparse.ArgumentParser(description="Save the parking site token to .env.")
    parser.add_argument("--refresh", action="store_true", help="headless re-read from the saved profile")
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    site_url = load_env().get("PARKING_URL")
    if not site_url:
        logger.error("❌ PARKING_URL missing in .env")
        return 2
    if not args.refresh:
        logger.info("ℹ️ A Chrome window will open. Sign in by hand; it closes by itself (10 min max).")
    token = grab_token(site_url, headless=args.refresh, wait_seconds=60 if args.refresh else 600)
    if not token:
        logger.error("❌ no valid token found - %s", "sign in did not complete in time"
                     if not args.refresh else "the saved profile is not logged in; run without --refresh")
        return 1
    set_env_value("PARKING_TOKEN", token)
    logger.info("✅ token saved; expires %s", auth.jwt_expiry(token).isoformat())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
