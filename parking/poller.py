"""One poll cycle: find uncovered target days, look for free slots, book them.

Run as ``python -m parking.poller`` (one cycle per invocation; the scheduler
provides the heartbeat). A day that already has an ACTIVE booking is never
booked again.
"""

from __future__ import annotations

import argparse
import logging
import os
import random
import sys
import time as time_module
from datetime import datetime, timedelta
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Callable, Dict, List, Optional
from zoneinfo import ZoneInfo

import requests

from parking import booking, burst, planner, report
from parking.api import ApiError, AuthError, ParkingApi, RateLimited, ServerError
from parking.booking import RunResult
from parking.config import (
    CONFIG_PATH, LOG_DIR, MESSAGE_LEDGER_PATH, STATE_PATH, Config, PollWindow, apply_env_overrides,
    load_config, load_env,
)
from parking.notify import FleetNotifier, Notifier
from parking.state import State, load_state, save_state

logger = logging.getLogger("parking.poller")

MAX_BACKOFF_MINUTES = 360


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


def _backoff(cfg: Config, state: State, now: datetime, minutes: Optional[int] = None) -> None:
    wait = minutes if minutes is not None else min(
        cfg.poll_minutes * (2 ** state.failures), MAX_BACKOFF_MINUTES)
    state.next_allowed = now + timedelta(minutes=wait)
    logger.warning("⚠️ backing off %s min (until %s)", wait, state.next_allowed.isoformat())


def _cycle(cfg: Config, api: ParkingApi, plate: str, now: datetime, notifier: Notifier,
           state: State, sleep: Callable[[float], None], rng: random.Random,
           say: Callable[[str], object], site_url: str = "") -> RunResult:
    pending = booking.pending_days(cfg, api.my_bookings(), now)
    if not pending:
        logger.info("ℹ️ every released target day is already booked")
        say("Every released target day already has a parking. Nothing to do.")
        return RunResult("nothing-pending")
    logger.info("ℹ️ %d target day(s) without a booking: %s", len(pending), pending)
    say(f"Checked my bookings: {len(pending)} day(s) still without a parking: {', '.join(pending)}.")
    say("Looking for free slots on the site...")

    slots = []
    for index, combo in enumerate(booking.combos(cfg)):
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
    outcomes = [booking.book_day(api, day, found[day], plate, cfg.type, cfg.dry_run, say)
                for day in pending if day in found]
    result.booked = [o.day for o in outcomes if o.status == "booked"]
    result.would_book = [o.day for o in outcomes if o.status == "would-book"]
    # Only now, with every day tried: screenshots between two bookings cost the 09-24 burst two days (#139).
    report.report_outcomes(outcomes, cfg, state, notifier, now, site_url)
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
        report.notify(notifier, cfg, env.get("PARKING_URL", ""), sweep_summary(result, now), disposable=True)
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

    say: Callable[[str], object] = ((lambda text: notifier.send(text, disposable=True))
                                    if is_verbose(cfg, now) else (lambda _text: None))
    say(f"Parking poll started ({now:%a %H:%M}).")

    token, stop = report.check_token(env, state, notifier, now, cfg.token_warning_days)
    if stop is not None:
        return stop

    endpoint = env.get("PARKING_API_URL", "")
    plate = env.get("PARKING_LICENSE_PLATE", "")
    if not endpoint or not plate:
        raise SystemExit("PARKING_API_URL and PARKING_LICENSE_PLATE must be set in .env")
    try:
        result = _cycle(cfg, api_factory(endpoint, token), plate, now, notifier, state, sleep, rng, say,
                        env.get("PARKING_URL", ""))
    except AuthError:
        logger.error("❌ token rejected (401/403)")
        report.alert(state, notifier, "login-needed",
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
    if result.status == "nothing-pending":
        # Every released day is booked, and nothing new can appear before the next release:
        # no site calls until then (a day cancelled by hand is one the owner no longer needs).
        state.next_allowed = planner.next_release(now, booking.release_time(cfg))
        logger.info("ℹ️ all covered: no site calls until the %s release", state.next_allowed.isoformat())
    else:
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
    parser.add_argument("--rehearse-burst", action="store_true",
                        help="run the Thursday burst now in dry-run, with its windows: books nothing, "
                             "no state or Telegram")
    args = parser.parse_args(argv)
    _setup_logging()
    logger.info("ℹ️ poller starting (pid=%d)", os.getpid())
    cfg = load_config(Path(args.config))
    env = load_env()
    cfg = apply_env_overrides(cfg, env)
    now_fn = lambda: datetime.now(ZoneInfo(cfg.timezone))  # noqa: E731
    if args.rehearse_burst:
        result = burst.rehearse(cfg, env, now_fn)
        logger.info("ℹ️ rehearsal finished: %s (would_book=%s)", result.status, result.would_book)
        return 1 if result.status == "error" else 0
    logger.info("ℹ️ mode: %s", "dry-run (no bookings)" if cfg.dry_run else "LIVE (will book)")

    state = load_state(STATE_PATH)
    now = now_fn()
    action = burst.burst_action(cfg, state, now)
    if action == "skip":
        # A burst outlasts the 5-minute tick, so the next tick can land while it
        # runs; it must not poll the site or post to (and clean up) the chat
        # alongside it, nor save a stale copy of the state over the burst's.
        logger.info("ℹ️ a burst is already running; nothing to do")
        return 0
    notifier = FleetNotifier(env, ledger_path=MESSAGE_LEDGER_PATH)
    try:
        if action == "burst":
            result = burst.run_burst(cfg, env, now, notifier, state, now_fn=now_fn)
        else:
            if cfg.jitter_max_seconds and not args.no_jitter:
                delay = random.uniform(0, cfg.jitter_max_seconds)
                logger.info("ℹ️ jitter: waiting %.0fs", delay)
                time_module.sleep(delay)
            now = now_fn()
            result = run_once(cfg, env, now, notifier, state)
    finally:
        notifier.close()
    save_state(STATE_PATH, state)
    logger.info("ℹ️ run finished: %s (booked=%s, would_book=%s)", result.status, result.booked,
                result.would_book)
    return 1 if result.status == "error" else 0


if __name__ == "__main__":
    raise SystemExit(main())
