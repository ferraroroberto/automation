"""One poll cycle: find uncovered target days, look for free slots, book them.

Run as ``python -m parking.poller`` (one cycle per invocation; the scheduler
provides the heartbeat). A day that already has an ACTIVE booking is never
booked again.
"""

from __future__ import annotations

import argparse
import logging
import random
import sys
import time as time_module
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Callable, Dict, List, Optional
from zoneinfo import ZoneInfo

import requests

from parking import auth, planner
from parking import screenshot as site_screenshot
from parking.api import ApiError, AuthError, ParkingApi, RateLimited, ServerError
from parking.config import (
    CONFIG_PATH, LOG_DIR, PROFILE_DIR, STATE_PATH, Config, PollWindow, apply_env_overrides, load_config, load_env,
)
from parking.notify import FleetNotifier, Notifier
from parking.state import State, load_state, save_state

logger = logging.getLogger("parking.poller")

MAX_BACKOFF_MINUTES = 360
ALERT_COOLDOWN = timedelta(hours=24)
SCREENSHOT_DIR = LOG_DIR / "screenshots"


@dataclass
class RunResult:
    status: str  # inactive | backoff | no-token | token-expired | nothing-pending | no-slots | checked | error
    pending: List[str] = field(default_factory=list)
    booked: List[str] = field(default_factory=list)
    would_book: List[str] = field(default_factory=list)
    error: str = ""


def _window(cfg: Config, now: datetime) -> Optional[PollWindow]:
    at = now.timetz().replace(tzinfo=None)
    for window in cfg.poll_windows:
        if window.days and now.weekday() not in window.days:
            continue
        wraps = window.start > window.end
        if (window.start <= at or at < window.end) if wraps else (window.start <= at < window.end):
            return window
    return None


def poll_interval(cfg: Config, now: datetime) -> int:
    """Minutes between polls right now: the first matching window, else `poll_minutes`. 0 = don't poll."""
    window = _window(cfg, now)
    return window.every_minutes if window else cfg.poll_minutes


def is_verbose(cfg: Config, now: datetime) -> bool:
    window = _window(cfg, now)
    return bool(window and window.verbose)


def _notify(notifier: Notifier, cfg: Config, site_url: str, text: str) -> bool:
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
                if notifier.send_file(text, str_paths[0]):
                    return True
            elif str_paths:
                if notifier.send_files(text, str_paths):
                    return True
        finally:
            for path in paths:
                path.unlink(missing_ok=True)
    return notifier.send(text)


def _alert(state: State, notifier: Notifier, key: str, text: str, now: datetime,
           cooldown: timedelta = ALERT_COOLDOWN, cfg: Optional[Config] = None,
           site_url: str = "") -> None:
    if state.alert_due(key, now, cooldown):
        if cfg is not None:
            _notify(notifier, cfg, site_url, text)
        else:
            notifier.send(text)
        state.mark_alert(key, now)


def _backoff(cfg: Config, state: State, now: datetime, minutes: Optional[int] = None) -> None:
    wait = minutes if minutes is not None else min(
        cfg.poll_minutes * (2 ** state.failures), MAX_BACKOFF_MINUTES)
    state.next_allowed = now + timedelta(minutes=wait)
    logger.warning("⚠️ backing off %s min (until %s)", wait, state.next_allowed.isoformat())


def _combos(cfg: Config) -> List[planner.Combo]:
    return [planner.Combo(c.id, c.name, s.id, s.name) for c in cfg.centers for s in cfg.sizes]


def _cycle(cfg: Config, api: ParkingApi, plate: str, now: datetime, notifier: Notifier,
           state: State, sleep: Callable[[float], None], rng: random.Random,
           say: Callable[[str], object], site_url: str = "") -> RunResult:
    bookings = api.my_bookings()
    covered, placeless = planner.covered_and_placeless(bookings, cfg.treat_placeless_as_covered)
    if placeless:
        logger.info("ℹ️ days with an ACTIVE booking but no place assigned: %s (covered=%s)",
                    sorted(placeless), cfg.treat_placeless_as_covered)
    holidays = planner.holiday_dates(cfg.holiday_country, cfg.holiday_subdiv, cfg.skip_dates,
                                     now.date(), cfg.horizon_days)
    all_targets = planner.target_dates(now.date(), cfg.weekdays, cfg.horizon_days)
    targets = [d for d in all_targets if d not in holidays]
    if len(targets) < len(all_targets):
        logger.info("ℹ️ skipping holiday(s): %s", [d for d in all_targets if d in holidays])
    pending = planner.pending_dates(targets, covered)
    if not pending:
        logger.info("ℹ️ every target day is already booked (%d checked)", len(targets))
        say("Every target day already has a parking. Nothing to do.")
        return RunResult("nothing-pending")
    logger.info("ℹ️ %d target day(s) without a booking: %s", len(pending), pending)
    say(f"Checked my bookings: {len(pending)} day(s) still without a parking: {', '.join(pending)}.")
    say("Looking for free slots on the site...")

    slots = []
    for index, combo in enumerate(_combos(cfg)):
        if index:
            sleep(rng.uniform(*cfg.request_pause_seconds))
        days = api.slot_days(combo.center_id, combo.size_id, cfg.type)
        logger.info("ℹ️ %s / %s: %d free day(s) %s", combo.center_name, combo.size_name, len(days),
                    sorted(planner.day_key(d) for d in days))
        slots.append((combo, days))
    found = planner.candidates_by_day(pending, slots)
    if not found:
        say("No free slot for those days yet.")
        return RunResult("no-slots", pending=pending)
    say(f"Free slot found for: {', '.join(sorted(found))}.")

    result = RunResult("checked", pending=pending)
    for day in pending:
        last_error: Optional[ApiError] = None
        for candidate in found.get(day, []):
            label = f"{day} at {candidate.combo.center_name} ({candidate.combo.size_name})"
            if cfg.dry_run:
                logger.info("ℹ️ [dry-run] would book %s", label)
                result.would_book.append(day)
                key = f"dry-run:{day}:{now.date().isoformat()}"
                _alert(state, notifier, key, f"[dry-run] a slot is free: {label}", now, timedelta(days=1))
                break
            say(f"Booking {label}...")
            try:
                api.create_booking(candidate.combo.center_id, candidate.combo.size_id, cfg.type,
                                   plate, candidate.raw_day)
            except ApiError as exc:
                if isinstance(exc, (AuthError, RateLimited, ServerError)):
                    raise
                logger.warning("⚠️ booking %s failed (%s); trying the next slot", label, exc)
                last_error = exc
                continue
            confirmed = planner.covered_and_placeless(api.my_bookings(), False)[0]
            if day in confirmed:
                logger.info("✅ booked %s", label)
                result.booked.append(day)
                _notify(notifier, cfg, site_url, f"Parking booked: {label}")
            else:
                logger.error("❌ createBooking returned but %s is not confirmed in my bookings", day)
                _notify(notifier, cfg, site_url,
                       f"Parking: booking {label} was sent but not confirmed - check the site")
            break
        else:
            if last_error is not None:
                _alert(state, notifier, f"book-failed:{day}",
                       f"Parking: a slot was free for {day} but booking failed ({last_error}). "
                       "Check the site.", now, timedelta(hours=1), cfg=cfg, site_url=site_url)
    return result


def sweep_summary(result: RunResult, now: datetime) -> str:
    """One line saying how a sweep ended, for the `sweep_report_until` trial."""
    head = f"Parking sweep {now:%a %H:%M}"
    if result.status == "nothing-pending":
        return f"{head} ✅ every target day already booked."
    if result.status in ("no-token", "token-expired"):
        return f"{head} ❌ login needed ({result.status}), site not checked."
    if result.status == "error":
        return f"{head} ❌ failed ({result.error}), will retry."
    open_days = f"{len(result.pending)} open day(s)"
    if result.booked:
        return f"{head} ✅ {open_days}, booked: {', '.join(result.booked)}."
    if result.would_book:
        return f"{head} ✅ {open_days}, free (dry-run): {', '.join(result.would_book)}."
    if result.status == "no-slots":
        return f"{head} ✅ {open_days}, no free slot for them."
    return f"{head} ⚠️ {open_days}, a slot was free but booking failed - check the site."


def run_once(cfg: Config, env: Dict[str, str], now: datetime, notifier: Notifier, state: State,
             api_factory: Callable[[str, str], ParkingApi] = ParkingApi,
             sleep: Callable[[float], None] = time_module.sleep,
             rng: Optional[random.Random] = None) -> RunResult:
    result = _run(cfg, env, now, notifier, state, api_factory, sleep, rng or random.Random())
    reporting = cfg.sweep_report_until is not None and now < cfg.sweep_report_until
    # Skipped runs are not sweeps; verbose windows already end with their own "Poll finished".
    if reporting and result.status not in ("backoff", "inactive") and not is_verbose(cfg, now):
        _notify(notifier, cfg, env.get("PARKING_URL", ""), sweep_summary(result, now))
    return result


def _run(cfg: Config, env: Dict[str, str], now: datetime, notifier: Notifier, state: State,
         api_factory: Callable[[str, str], ParkingApi], sleep: Callable[[float], None],
         rng: random.Random) -> RunResult:
    if state.in_backoff(now):
        logger.info("ℹ️ not due until %s", state.next_allowed.isoformat())
        return RunResult("backoff")
    interval = poll_interval(cfg, now)
    if not interval:
        logger.info("ℹ️ polling is off in this time window")
        return RunResult("inactive")

    say = notifier.send if is_verbose(cfg, now) else (lambda _text: None)
    say(f"Parking poll started ({now:%a %H:%M}).")

    token = auth.normalize_token(env.get("PARKING_TOKEN"))
    if not token:
        _alert(state, notifier, "login-needed",
               "Parking: login needed (no token). Run: python -m parking.login", now)
        return RunResult("no-token")
    left = auth.time_left(token, now)
    if left is not None and left <= timedelta(0):
        _alert(state, notifier, "login-needed",
               "Parking: the login token has expired. Run: python -m parking.login", now)
        return RunResult("token-expired")
    if left is not None and left <= timedelta(days=cfg.token_warning_days):
        _alert(state, notifier, "token-expiring",
               f"Parking: login token expires in {left.days}d {left.seconds // 3600}h. "
               "Run: python -m parking.login", now)

    endpoint = env.get("PARKING_API_URL", "")
    plate = env.get("PARKING_LICENSE_PLATE", "")
    if not endpoint or not plate:
        raise SystemExit("PARKING_API_URL and PARKING_LICENSE_PLATE must be set in .env")
    try:
        result = _cycle(cfg, api_factory(endpoint, token), plate, now, notifier, state, sleep, rng, say,
                        env.get("PARKING_URL", ""))
    except AuthError:
        logger.error("❌ token rejected (401/403)")
        _alert(state, notifier, "login-needed",
               "Parking: the site rejected the token. Run: python -m parking.login", now)
        _backoff(cfg, state, now, MAX_BACKOFF_MINUTES)
        return RunResult("error", error="token rejected")
    except (RateLimited, ServerError, requests.RequestException, ApiError) as exc:
        state.failures += 1
        logger.error("❌ %s (failure #%d)", type(exc).__name__, state.failures)
        if state.failures == cfg.repeated_error_threshold:
            notifier.send(f"Parking: {state.failures} consecutive errors ({type(exc).__name__}); backing off")
        _backoff(cfg, state, now)
        say(f"Poll failed ({type(exc).__name__}); will retry.")
        return RunResult("error", error=type(exc).__name__)
    state.failures = 0
    # The scheduler fires at the finest interval; later runs exit at the in_backoff check above.
    # Minus the start jitter, so a run that jitters earlier than the last one is not skipped.
    state.next_allowed = now + timedelta(minutes=interval, seconds=-cfg.jitter_max_seconds)
    say(f"Poll finished: {result.status}, booked {len(result.booked)} day(s).")
    return result


def _setup_logging() -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s")
    handlers = [RotatingFileHandler(LOG_DIR / "poller.log", maxBytes=1_000_000, backupCount=3,
                                    encoding="utf-8"),
                logging.StreamHandler(sys.stdout)]
    for handler in handlers:
        handler.setFormatter(fmt)
    logging.basicConfig(level=logging.INFO, handlers=handlers)
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Run one parking poll cycle.")
    parser.add_argument("--config", default=str(CONFIG_PATH))
    parser.add_argument("--no-jitter", action="store_true", help="skip the random start delay")
    args = parser.parse_args(argv)
    _setup_logging()
    cfg = load_config(Path(args.config))
    env = load_env()
    cfg = apply_env_overrides(cfg, env)
    logger.info("ℹ️ mode: %s", "dry-run (no bookings)" if cfg.dry_run else "LIVE (will book)")
    if cfg.jitter_max_seconds and not args.no_jitter:
        delay = random.uniform(0, cfg.jitter_max_seconds)
        logger.info("ℹ️ jitter: waiting %.0fs", delay)
        time_module.sleep(delay)
    now = datetime.now(ZoneInfo(cfg.timezone))
    state = load_state(STATE_PATH)
    result = run_once(cfg, env, now, FleetNotifier(env), state)
    save_state(STATE_PATH, state)
    logger.info("ℹ️ run finished: %s (booked=%s, would_book=%s)", result.status, result.booked,
                result.would_book)
    return 1 if result.status == "error" else 0


if __name__ == "__main__":
    raise SystemExit(main())
