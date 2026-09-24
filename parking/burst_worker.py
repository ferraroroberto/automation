"""One day's Thursday burst worker, in its own console window (started by `parking.burst`).

``python -m parking.burst_worker --day 2026-10-05 --target <iso> --deadline <iso> --run-dir <dir> [--dry-run]``

Waits precisely for `target`, then scans the centres/sizes in config order and
books the first free one for its day; a REJECTED answer moves on to the next,
and when nothing sticks it waits `poll_interval_seconds` and scans again - until
the day is ACTIVE or `deadline` passes. Progress goes to this console, to
``logs/burst/<day>.log`` and, as JSON lines, to ``<run-dir>/<day>.jsonl`` for the
orchestrator; the last line (``"done": true``) is the result.

Never books a day that already has an ACTIVE booking (checked before the
target), never more than one booking, and in dry-run never calls createBooking.
"""

from __future__ import annotations

import argparse
import ctypes
import json
import logging
import random
import sys
import time as time_module
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable, List, Optional, Set

import requests

from parking import auth, booking, planner
from parking.api import ApiError, ParkingApi
from parking.booking import DayOutcome
from parking.burst import sleep_until_precise
from parking.config import LOG_DIR, Config, apply_env_overrides, load_config, load_env

logger = logging.getLogger("parking.burst_worker")

SCAN_PAUSE_SECONDS = (0.05, 0.15)  # between the combos of one scan
HOLD_SECONDS = 120  # the console stays open this long after the result, for the owner to read


def _scan_and_book(api: ParkingApi, cfg: Config, day: str, plate: str, dry_run: bool,
                   sleep: Callable[[float], None], rng: random.Random, refused: Set[planner.Combo]) -> DayOutcome:
    """One pass over the combos, freshest data first: book `day` at the first one free.

    A combo whose booking failed joins `refused` and is skipped for the rest of the
    burst: every REJECTED attempt leaves a record on the site, so a slot that is
    listed free but always refused must not be retried every check.
    """
    rejected: List[str] = []
    for index, combo in enumerate(c for c in booking.combos(cfg) if c not in refused):
        if index:
            sleep(rng.uniform(*SCAN_PAUSE_SECONDS))
        free = [raw for raw in api.slot_days(combo.center_id, combo.size_id, cfg.type)
                if planner.day_key(raw) == day]
        if not free:
            continue
        logger.info("ℹ️ %s / %s: free", combo.center_name, combo.size_name)
        outcome = booking.book_day(api, day, [planner.Candidate(day, free[0], combo)], plate, cfg.type,
                                   dry_run)
        if outcome.status != "failed":
            return outcome
        refused.add(combo)
        rejected.append(outcome.error)
    return DayOutcome(day, "failed", error="; ".join(rejected)) if rejected else DayOutcome(day, "no-slot")


def work_day(api: ParkingApi, cfg: Config, day: str, plate: str, target: datetime, deadline: datetime,
             dry_run: bool, emit: Callable[[str], None], sleep: Callable[[float], None],
             now_fn: Callable[[], datetime], rng: Optional[random.Random] = None) -> DayOutcome:
    """Book `day` from `target` until it sticks or `deadline` passes (see the module docstring)."""
    rng = rng or random.Random()
    burst = cfg.thursday_burst
    interval = burst.poll_interval_seconds if burst else 1.5
    # Before the target: the hard rule check, and it opens the connection so 16:00 starts warm.
    covered, _ = planner.covered_and_placeless(api.my_bookings(), cfg.treat_placeless_as_covered)
    if day in covered:
        if not dry_run:
            logger.info("ℹ️ %s already has an ACTIVE booking; nothing to do", day)
            return DayOutcome(day, "covered")
        logger.info("ℹ️ %s already has a booking - dry-run, so scanning anyway (nothing is booked)", day)
    logger.info("⚡ waiting for %s", target.isoformat())
    sleep_until_precise(target, sleep, now_fn)
    checks = 0
    refused: Set[planner.Combo] = set()
    while True:
        checks += 1
        at = now_fn()
        outcome = _scan_and_book(api, cfg, day, plate, dry_run, sleep, rng, refused)
        line = {"no-slot": "nothing free", "booked": f"✅ booked {outcome.label}",
                "would-book": f"🟢 free (dry-run): {outcome.label}",
                "failed": f"⚠️ {outcome.error}", "unconfirmed": f"⚠️ sent, not confirmed: {outcome.label}"}
        emit(f"#{checks} {at:%H:%M:%S}.{at.microsecond // 10000:02d} {line[outcome.status]}")
        if outcome.status in ("booked", "would-book", "unconfirmed"):
            return outcome
        if now_fn() + timedelta(seconds=interval) >= deadline:
            return outcome
        sleep(interval)


def _set_title(text: str) -> None:
    if sys.platform == "win32":
        ctypes.windll.kernel32.SetConsoleTitleW(text)


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="One day's Thursday burst worker.")
    parser.add_argument("--day", required=True)
    parser.add_argument("--target", required=True, type=datetime.fromisoformat)
    parser.add_argument("--deadline", required=True, type=datetime.fromisoformat)
    parser.add_argument("--run-dir", required=True, type=Path)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--hold-seconds", type=float, default=HOLD_SECONDS)
    args = parser.parse_args(argv)

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    log_dir = LOG_DIR / "burst"
    log_dir.mkdir(parents=True, exist_ok=True)
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    # Its own file (overwritten each burst): RotatingFileHandler is not safe across processes.
    handlers = [logging.FileHandler(log_dir / f"{args.day}.log", mode="w", encoding="utf-8"),
                logging.StreamHandler(sys.stdout)]
    for handler in handlers:
        handler.setFormatter(fmt)
    logging.basicConfig(level=logging.INFO, handlers=handlers)

    env = load_env()
    cfg = apply_env_overrides(load_config(), env)
    dry_run = args.dry_run or cfg.dry_run
    _set_title(f"Parking burst {args.day}{' (dry-run)' if dry_run else ''}")
    logger.info("ℹ️ worker for %s (%s)", args.day, "dry-run, books nothing" if dry_run else "LIVE, will book")

    events = args.run_dir / f"{args.day}.jsonl"

    def write(record: dict) -> None:
        with events.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(record) + "\n")

    def emit(text: str) -> None:
        logger.info("⚡ %s", text)
        write({"text": text})

    token = auth.normalize_token(env.get("PARKING_TOKEN"))
    try:
        api = ParkingApi(env.get("PARKING_API_URL", ""), token or "")
        outcome = work_day(api, cfg, args.day, env.get("PARKING_LICENSE_PLATE", ""), args.target,
                           args.deadline, dry_run, emit, time_module.sleep,
                           lambda: datetime.now(args.target.tzinfo))
    except (requests.RequestException, ApiError) as exc:
        # Auth, rate limit or server trouble: stop this day rather than push through.
        logger.error("❌ %s: %s - stopping", type(exc).__name__, exc)
        outcome = DayOutcome(args.day, "error", error=type(exc).__name__)
    write({"done": True, "status": outcome.status, "label": outcome.label, "error": outcome.error})
    logger.info("🏁 %s: %s %s", args.day, outcome.status, outcome.label or outcome.error)
    logger.info("ℹ️ this window closes in %.0fs", args.hold_seconds)
    time_module.sleep(args.hold_seconds)
    return 0 if outcome.status != "error" else 1


if __name__ == "__main__":
    raise SystemExit(main())
