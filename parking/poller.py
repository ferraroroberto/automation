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
from parking.api import ApiError, AuthError, ParkingApi, RateLimited, ServerError
from parking.config import (
    CONFIG_PATH, LOG_DIR, STATE_PATH, Config, apply_env_overrides, load_config, load_env,
)
from parking.notify import FleetNotifier, Notifier
from parking.state import State, load_state, save_state

logger = logging.getLogger("parking.poller")

MAX_BACKOFF_MINUTES = 360
ALERT_COOLDOWN = timedelta(hours=24)


@dataclass
class RunResult:
    status: str  # inactive | backoff | no-token | token-expired | nothing-pending | no-slots | checked | error
    pending: List[str] = field(default_factory=list)
    booked: List[str] = field(default_factory=list)
    would_book: List[str] = field(default_factory=list)


def in_active_hours(cfg: Config, now: datetime) -> bool:
    return cfg.active_start <= now.timetz().replace(tzinfo=None) <= cfg.active_end


def _alert(state: State, notifier: Notifier, key: str, text: str, now: datetime,
           cooldown: timedelta = ALERT_COOLDOWN) -> None:
    if state.alert_due(key, now, cooldown):
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
           state: State, sleep: Callable[[float], None], rng: random.Random) -> RunResult:
    bookings = api.my_bookings()
    covered, placeless = planner.covered_and_placeless(bookings, cfg.treat_placeless_as_covered)
    if placeless:
        logger.info("ℹ️ days with an ACTIVE booking but no place assigned: %s (covered=%s)",
                    sorted(placeless), cfg.treat_placeless_as_covered)
    targets = planner.target_dates(now.date(), cfg.weekdays, cfg.horizon_days)
    pending = planner.pending_dates(targets, covered)
    if not pending:
        logger.info("ℹ️ every target day is already booked (%d checked)", len(targets))
        return RunResult("nothing-pending")
    logger.info("ℹ️ %d target day(s) without a booking: %s", len(pending), pending)

    slots = []
    for index, combo in enumerate(_combos(cfg)):
        if index:
            sleep(rng.uniform(*cfg.request_pause_seconds))
        days = api.slot_days(combo.center_id, combo.size_id, cfg.type)
        logger.info("ℹ️ %s / %s: %d free day(s)", combo.center_name, combo.size_name, len(days))
        slots.append((combo, days))
    found = planner.candidates_by_day(pending, slots)
    if not found:
        return RunResult("no-slots", pending=pending)

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
                notifier.send(f"Parking booked: {label}")
            else:
                logger.error("❌ createBooking returned but %s is not confirmed in my bookings", day)
                notifier.send(f"Parking: booking {label} was sent but not confirmed - check the site")
            break
        else:
            if last_error is not None:
                _alert(state, notifier, f"book-failed:{day}",
                       f"Parking: a slot was free for {day} but booking failed ({last_error}). "
                       "Check the site.", now, timedelta(hours=1))
    return result


def run_once(cfg: Config, env: Dict[str, str], now: datetime, notifier: Notifier, state: State,
             api_factory: Callable[[str, str], ParkingApi] = ParkingApi,
             sleep: Callable[[float], None] = time_module.sleep,
             rng: Optional[random.Random] = None) -> RunResult:
    rng = rng or random.Random()
    if state.in_backoff(now):
        logger.info("ℹ️ in backoff until %s", state.next_allowed.isoformat())
        return RunResult("backoff")
    if not in_active_hours(cfg, now):
        logger.info("ℹ️ outside active hours")
        return RunResult("inactive")

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
        result = _cycle(cfg, api_factory(endpoint, token), plate, now, notifier, state, sleep, rng)
    except AuthError:
        logger.error("❌ token rejected (401/403)")
        _alert(state, notifier, "login-needed",
               "Parking: the site rejected the token. Run: python -m parking.login", now)
        _backoff(cfg, state, now, MAX_BACKOFF_MINUTES)
        return RunResult("error")
    except (RateLimited, ServerError, requests.RequestException, ApiError) as exc:
        state.failures += 1
        logger.error("❌ %s (failure #%d)", type(exc).__name__, state.failures)
        if state.failures == cfg.repeated_error_threshold:
            notifier.send(f"Parking: {state.failures} consecutive errors ({type(exc).__name__}); backing off")
        _backoff(cfg, state, now)
        return RunResult("error")
    state.failures = 0
    state.next_allowed = None
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
