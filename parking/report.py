"""Telegram notifications for poll/burst outcomes, with the calendar screenshots attached.

Shared by the regular poll and the Thursday burst. Callers send these only
after every day has been tried: capturing screenshots between two bookings
cost the 2026-09-24 burst two days (#139).
"""

from __future__ import annotations

import logging
from datetime import datetime, timedelta
from typing import Dict, Iterable, Optional, Tuple

from parking import auth
from parking import screenshot as site_screenshot
from parking.booking import DayOutcome, RunResult
from parking.config import LOG_DIR, PROFILE_DIR, Config
from parking.notify import Notifier
from parking.state import State

logger = logging.getLogger("parking.report")

ALERT_COOLDOWN = timedelta(hours=24)
SCREENSHOT_DIR = LOG_DIR / "screenshots"


def notify(notifier: Notifier, cfg: Config, site_url: str, text: str, disposable: bool = False) -> bool:
    """Send `text`, attaching each center's calendar (this month + next) as one
    Telegram message when screenshots can be captured.

    Best-effort throughout: a missing `PARKING_URL`, no screenshots captured,
    or the send itself failing all fall back to a plain text notification.
    Whatever was captured is deleted after the send attempt either way, so
    screenshots never accumulate between polls.
    """
    if site_url:
        paths = site_screenshot.capture_all([c.name for c in cfg.centers], site_url,
                                            SCREENSHOT_DIR, PROFILE_DIR)
        try:
            str_paths = [str(p) for p in paths]
            if len(str_paths) == 1:
                if notifier.send_file(text, str_paths[0], disposable=disposable):
                    return True
            elif str_paths:
                if notifier.send_files(text, str_paths, disposable=disposable):
                    return True
        finally:
            for path in paths:
                path.unlink(missing_ok=True)
    return notifier.send(text, disposable=disposable)


def alert(state: State, notifier: Notifier, key: str, text: str, now: datetime,
          cooldown: timedelta = ALERT_COOLDOWN, cfg: Optional[Config] = None,
          site_url: str = "") -> None:
    if state.alert_due(key, now, cooldown):
        if cfg is not None:
            notify(notifier, cfg, site_url, text)
        else:
            notifier.send(text)
        state.mark_alert(key, now)


def report_outcomes(outcomes: Iterable[DayOutcome], cfg: Config, state: State, notifier: Notifier,
                    now: datetime, site_url: str) -> None:
    """One message for every day booked, then one per unconfirmed / failed / broken day
    (dry-run: one alert per free day). Each message with screenshots costs a browser
    capture - slow while the site is busy - so the bookings share one."""
    outcomes = list(outcomes)
    booked = [o.label for o in outcomes if o.status == "booked"]
    if booked:
        notify(notifier, cfg, site_url, "Parking booked: " + "; ".join(booked))
    for outcome in outcomes:
        if outcome.status == "would-book":
            alert(state, notifier, f"dry-run:{outcome.day}:{now.date().isoformat()}",
                  f"[dry-run] a slot is free: {outcome.label}", now, timedelta(days=1))
        elif outcome.status == "unconfirmed":
            notify(notifier, cfg, site_url,
                   f"Parking: booking {outcome.label} was sent but not confirmed - check the site")
        elif outcome.status == "failed":
            alert(state, notifier, f"book-failed:{outcome.day}",
                  f"Parking: a slot was free for {outcome.day} but booking failed ({outcome.error}). "
                  "Check the site.", now, timedelta(hours=1), cfg=cfg, site_url=site_url)
        elif outcome.status in ("error", "unknown"):
            notifier.send(f"Parking burst: {outcome.day} ended {outcome.status} ({outcome.error}) - check the site")


def check_token(env: Dict[str, str], state: State, notifier: Notifier, now: datetime,
                warning_days: Optional[int] = None) -> Tuple[str, Optional[RunResult]]:
    """The site token, or the RunResult to stop with (after a once-a-day login alert).

    With `warning_days`, a token that expires within them also alerts, but the run goes on.
    """
    token = auth.normalize_token(env.get("PARKING_TOKEN"))
    if not token:
        alert(state, notifier, "login-needed",
              "Parking: login needed (no token). Run: python -m parking.login", now)
        return "", RunResult("no-token")
    left = auth.time_left(token, now)
    if left is not None and left <= timedelta(0):
        alert(state, notifier, "login-needed",
              "Parking: the login token has expired. Run: python -m parking.login", now)
        return "", RunResult("token-expired")
    if warning_days is not None and left is not None and left <= timedelta(days=warning_days):
        alert(state, notifier, "token-expiring",
              f"Parking: login token expires in {left.days}d {left.seconds // 3600}h. "
              "Run: python -m parking.login", now)
    return token, None
