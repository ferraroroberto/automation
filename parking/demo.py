"""Rehearsal of the Thursday 16:00 poll for whoever reads the Telegram chat.

Run: ``python -m parking.demo``. It drives the real ``run_once`` against a fake site that has
free slots for a few days, and sends the live-log messages to the real chat, each prefixed
``[DEMO]``. The fake API cannot reach the real site, so nothing is ever booked.
"""

from __future__ import annotations

import base64
import json
import logging
import random
import sys
import time as time_module
from dataclasses import replace
from datetime import datetime, timedelta
from typing import Callable, Dict, List, Optional
from zoneinfo import ZoneInfo

from parking import planner, poller
from parking.config import Config, apply_env_overrides, load_config, load_env
from parking.notify import FleetNotifier, Notifier
from parking.state import State

logger = logging.getLogger("parking.demo")

DEMO_DAYS = 3
MESSAGE_PAUSE_SECONDS = 2.0


class DemoNotifier:
    """Prefixes every message with [DEMO] and paces them so the chat reads like a live log."""

    def __init__(self, inner: Notifier, pause: Callable[[float], None] = time_module.sleep) -> None:
        self._inner = inner
        self._pause = pause

    def send(self, text: str, disposable: bool = False) -> bool:
        sent = self._inner.send(f"[DEMO] {text}", disposable=disposable)
        self._pause(MESSAGE_PAUSE_SECONDS)
        return sent

    def send_file(self, text: str, path: str, disposable: bool = False) -> bool:
        sent = self._inner.send_file(f"[DEMO] {text}", path, disposable=disposable)
        self._pause(MESSAGE_PAUSE_SECONDS)
        return sent

    def send_files(self, text: str, paths: List[str], disposable: bool = False) -> bool:
        sent = self._inner.send_files(f"[DEMO] {text}", paths, disposable=disposable)
        self._pause(MESSAGE_PAUSE_SECONDS)
        return sent

    def ping(self, text: str) -> None:
        self._inner.ping(f"[DEMO] {text}")

    def close(self) -> None:
        self._inner.close()


class DemoApi:
    """A site with one free slot (22@ / Large) on each of `days`; booking just records it."""

    def __init__(self, days: List[str]) -> None:
        self._free = [f"{day}T10:00:00.000Z" for day in days]
        self._bookings: List[Dict] = []

    def my_bookings(self) -> List[Dict]:
        return list(self._bookings)

    def slot_days(self, center_id: int, size_id: int, type_: str) -> List[str]:
        return list(self._free) if (center_id, size_id) == (1, 3) else []

    def create_booking(self, center_id: int, size_id: int, type_: str, plate: str, raw_day: str) -> List:
        self._bookings.append({"day": raw_day, "state": "ACTIVE", "place": {"place": "DEMO"}})
        return []


def _fake_token(now: datetime) -> str:
    def enc(obj: dict) -> str:
        return base64.urlsafe_b64encode(json.dumps(obj).encode()).decode().rstrip("=")

    return f"{enc({'alg': 'none'})}.{enc({'exp': int((now + timedelta(days=20)).timestamp())})}.demo"


def next_thursday_poll(cfg: Config, today: datetime) -> datetime:
    ahead = (3 - today.weekday()) % 7 or 7
    return (today + timedelta(days=ahead)).replace(hour=16, minute=2, second=0, microsecond=0)


def run_demo(cfg: Config, notifier: Notifier, today: Optional[datetime] = None) -> poller.RunResult:
    today = today or datetime.now(ZoneInfo(cfg.timezone))
    now = next_thursday_poll(cfg, today)
    skip = planner.holiday_dates(cfg.holiday_country, cfg.holiday_subdiv, cfg.skip_dates,
                                 now.date(), cfg.horizon_days)
    days = [d for d in planner.target_dates(now.date(), cfg.weekdays, cfg.horizon_days) if d not in skip]
    api = DemoApi(days[:DEMO_DAYS])
    env = {"PARKING_TOKEN": _fake_token(now), "PARKING_API_URL": "demo", "PARKING_LICENSE_PLATE": "DEMO"}
    demo_cfg = replace(cfg, dry_run=False, jitter_max_seconds=0)
    notifier.send("Rehearsal of the Thursday 16:00 poll. Nothing real is booked, this is a simulation.")
    return poller.run_once(demo_cfg, env, now, notifier, State(), api_factory=lambda *_: api,
                           sleep=lambda _s: None, rng=random.Random(0))


def main() -> int:
    logging.basicConfig(level=logging.INFO)
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    cfg = apply_env_overrides(load_config(), load_env())
    result = run_demo(cfg, DemoNotifier(FleetNotifier(load_env())))
    logger.info("ℹ️ demo finished: %s", result.status)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
