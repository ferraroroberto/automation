"""Shared human-looking Chrome launch for Playwright automations against third-party sites.

Fleet rule "Browser automation must not look like a bot": real Chrome (``channel="chrome"``),
a persistent profile, 1280x900, ``navigator.webdriver`` undefined, the automation infobar and
``--no-sandbox`` warning stripped. A persistent profile allows one live Chrome at a time, so a
busy profile is waited for with backoff (60/120/240/480 s), never killed.

Scripts here are run as ``python -m <folder>.<module>`` from the repo root, so this module
imports directly; a script run as ``python <folder>/<script>.py`` appends the repo root to
``sys.path`` first (see ``no_window.py``).
"""

from __future__ import annotations

import logging
import re
import time
from typing import Callable, Iterable, Sequence

logger = logging.getLogger("browser_stealth")

STEALTH_INIT_SCRIPT = "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
PROFILE_BUSY_DELAYS: tuple[int, ...] = (60, 120, 240, 480)

# Chrome exits with code 21 (headless) or hands off to the running instance (headed) when the
# user-data-dir is already open; Playwright then reports the target as closed.
_PROFILE_BUSY_RE = re.compile(r"exitCode=21|existing browser session|already in use", re.IGNORECASE)


class ProfileBusy(RuntimeError):
    """The persistent profile stayed in use by another Chrome for the whole backoff schedule."""


def stealth_launch_kwargs(user_data_dir: str, headless: bool,
                          disabled_features: Iterable[str] = ()) -> dict:
    """``launch_persistent_context`` kwargs; ``disabled_features`` join Translate in one flag."""
    features = ",".join(["Translate", *disabled_features])
    return {
        "user_data_dir": user_data_dir,
        "channel": "chrome",
        "headless": headless,
        "locale": "en-US",
        "args": [
            "--disable-blink-features=AutomationControlled",
            f"--disable-features={features}",
            "--no-default-browser-check",
            "--no-first-run",
            "--lang=en-US",
        ],
        "ignore_default_args": ["--enable-automation", "--enable-blink-features=IdleDetection"],
        "viewport": {"width": 1280, "height": 900},
        "chromium_sandbox": True,
    }


def is_profile_busy(exc: BaseException) -> bool:
    return bool(_PROFILE_BUSY_RE.search(str(exc)))


def launch_persistent(playwright, user_data_dir: str, headless: bool,
                      disabled_features: Iterable[str] = (),
                      delays: Sequence[int] = PROFILE_BUSY_DELAYS,
                      sleep: Callable[[float], None] = time.sleep):
    """Stealth persistent context with the init script applied; waits out a busy profile."""
    kwargs = stealth_launch_kwargs(user_data_dir, headless, disabled_features)
    for attempt in range(len(delays) + 1):
        try:
            context = playwright.chromium.launch_persistent_context(**kwargs)
        except Exception as exc:
            if not is_profile_busy(exc):
                raise
            if attempt == len(delays):
                raise ProfileBusy(
                    f"Chrome profile {user_data_dir} is still in use after waiting "
                    f"{sum(delays)} s; close the other Chrome window using it") from exc
            logger.warning("⚠️ Chrome profile in use by another window; retrying in %d s", delays[attempt])
            sleep(delays[attempt])
            continue
        context.add_init_script(STEALTH_INIT_SCRIPT)
        return context
    raise AssertionError("unreachable")
