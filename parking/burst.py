"""The Thursday 16:00 burst: one visible worker process per day, watched in Chrome.

The tick that lands in the watch band before `target` becomes the orchestrator.
Before the target it:
- reads my bookings and works out the days released at the target,
- opens one headed observer window per day (`parking.observer`),
- starts one worker per day (``python -m parking.burst_worker``), each in its own console.

Each worker waits for the target itself, then scans and books its own day until
it sticks or the burst window ends, so one day's slow step or rejection never
delays another (the 2026-09-24 burst lost two days that way, #139). The
orchestrator relays the workers' progress to Telegram as batched pings and,
once every worker has a result, sends the booking confirmations with screenshots.

``python -m parking.poller --rehearse-burst`` runs the same flow a few seconds
from now in forced dry-run: nothing booked, no state written, nothing sent.
"""

from __future__ import annotations

import json
import logging
import subprocess
import sys
import time as time_module
from dataclasses import replace
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

import requests

from parking import booking, planner, report
from parking.api import ApiError, AuthError, ParkingApi, RateLimited, ServerError
from parking.booking import DayOutcome, RunResult
from parking.config import PROFILE_DIR, ROOT, STATE_PATH, Config, ThursdayBurst
from parking.notify import LogNotifier, Notifier
from parking.observer import Observers
from parking.planner import WEEKDAYS
from parking.state import State, save_state

logger = logging.getLogger("parking.burst")

BURST_STATE_KEY = "thursday-burst"
BURST_COOLDOWN = timedelta(hours=20)  # at most once per Thursday
FINE_APPROACH_SECONDS = 2.0  # last stretch of sleep_until_precise: busy-poll instead of one long sleep
FINE_POLL_SECONDS = 0.02

RUN_DIR = STATE_PATH.parent / "burst"  # per-day event files + the running marker
RUNNING_MAX_AGE = timedelta(minutes=20)  # a marker older than this is from a killed run
WORKER_GRACE_SECONDS = 60  # past the deadline: a last in-flight request (20 s timeout) still lands
FOLLOW_POLL_SECONDS = 0.25
OBSERVER_HOLD_SECONDS = 120  # windows stay up after the last result, for the owner to look
REHEARSAL_LEAD_SECONDS = 20  # time to open the windows before the rehearsal's "16:00"
REHEARSAL_DURATION_SECONDS = 30

CREATE_NEW_CONSOLE = getattr(subprocess, "CREATE_NEW_CONSOLE", 0)


def burst_target(burst: ThursdayBurst, now: datetime) -> datetime:
    return now.replace(hour=burst.target.hour, minute=burst.target.minute, second=0, microsecond=0)


def in_watch_band(burst: ThursdayBurst, now: datetime) -> bool:
    """True from `watch_before_minutes` before `target` through the end of the burst window.

    Any regular invocation landing in here is the one that arms the burst -
    it doesn't matter which of the day's ticks it is, since the burst itself
    self-times a precise wait to `target` regardless of when it started.
    """
    if now.weekday() != WEEKDAYS.index("thu") and now.date() not in burst.trial_dates:
        return False
    target = burst_target(burst, now)
    start = target - timedelta(minutes=burst.watch_before_minutes)
    end = target + timedelta(seconds=burst.duration_seconds)
    return start <= now <= end


def is_running(now: datetime, run_dir: Optional[Path] = None) -> bool:
    """A burst (or rehearsal) is in progress: its marker exists and isn't stale."""
    try:
        marker = (run_dir or RUN_DIR) / "running.json"
        started = datetime.fromisoformat(json.loads(marker.read_text())["started"])
    except (OSError, ValueError, KeyError):
        return False
    return now - started < RUNNING_MAX_AGE


def burst_action(cfg: Config, state: State, now: datetime, run_dir: Optional[Path] = None) -> str:
    """"burst" to arm the burst now, "skip" while one is armed or running, else "normal"."""
    if is_running(now, run_dir):
        return "skip"
    burst = cfg.thursday_burst
    if not (burst and burst.enabled and in_watch_band(burst, now)):
        return "normal"
    return "burst" if state.alert_due(BURST_STATE_KEY, now, BURST_COOLDOWN) else "skip"


def sleep_until_precise(target: datetime, sleep: Callable[[float], None],
                        now_fn: Callable[[], datetime]) -> None:
    """Sleep until `target`: one coarse sleep, then a short busy-poll for the
    final stretch, so a single long `time.sleep()`'s OS timer resolution
    doesn't decide how close to the second we actually land.
    """
    remaining = (target - now_fn()).total_seconds()
    if remaining > FINE_APPROACH_SECONDS:
        sleep(remaining - FINE_APPROACH_SECONDS)
    while now_fn() < target:
        sleep(FINE_POLL_SECONDS)


def spawn_worker(day: str, target: datetime, deadline: datetime, dry_run: bool,
                 run_dir: Path) -> subprocess.Popen:
    """One worker in its own console window (meant to be seen, hence no CREATE_NO_WINDOW)."""
    argv = [sys.executable, "-m", "parking.burst_worker", "--day", day, "--target", target.isoformat(),
            "--deadline", deadline.isoformat(), "--run-dir", str(run_dir)]
    if dry_run:
        argv.append("--dry-run")
    return subprocess.Popen(argv, cwd=str(ROOT.parent), creationflags=CREATE_NEW_CONSOLE)


def _short(day: str) -> str:
    return f"{day[8:10]}/{day[5:7]}"


def _read_events(path: Path, offset: int) -> Tuple[List[dict], int]:
    """New complete JSON lines in `path` from byte `offset`: (events, new offset)."""
    try:
        with path.open("rb") as handle:
            handle.seek(offset)
            chunk = handle.read()
    except FileNotFoundError:
        return [], offset
    complete = chunk[: chunk.rfind(b"\n") + 1]  # a line still being written waits for the next read
    events = []
    for line in complete.decode("utf-8").splitlines():
        try:
            events.append(json.loads(line))
        except ValueError:
            logger.warning("⚠️ unreadable worker event: %r", line)
    return events, offset + len(complete)


def follow_workers(workers: Dict[str, "subprocess.Popen"], run_dir: Path, until: datetime,
                   notifier: Notifier, observers: Optional[Observers], sleep: Callable[[float], None],
                   now_fn: Callable[[], datetime]) -> Dict[str, DayOutcome]:
    """Relay each worker's progress as pings until every one has reported its result.

    A worker that exits without a result is an "error"; one still silent at
    `until` is "unknown" - never counted as booked or as nothing-free.
    """
    offsets = {day: 0 for day in workers}
    outcomes: Dict[str, DayOutcome] = {}
    while True:
        for day, worker in workers.items():
            if day in outcomes:
                continue
            events, offsets[day] = _read_events(run_dir / f"{day}.jsonl", offsets[day])
            for event in events:
                if event.get("done"):
                    outcomes[day] = DayOutcome(day, event["status"], event.get("label", ""),
                                               event.get("error", ""))
                    logger.info("ℹ️ worker %s finished: %s %s", day, event["status"],
                                event.get("label") or event.get("error", ""))
                    if observers is not None:
                        observers.refresh(day)
                else:
                    logger.info("ℹ️ worker %s: %s", day, event.get("text", ""))
                    notifier.ping(f"⚡ {_short(day)} {event.get('text', '')}")
            if day not in outcomes and worker.poll() is not None:
                outcomes[day] = DayOutcome(day, "error", error=f"worker exited ({worker.returncode}) "
                                                               "without a result")
                logger.error("❌ worker %s exited (%s) without a result", day, worker.returncode)
        if len(outcomes) == len(workers):
            return outcomes
        if now_fn() > until:
            for day in workers:
                if day not in outcomes:
                    outcomes[day] = DayOutcome(day, "unknown", error="no result from its worker in time")
                    logger.error("❌ worker %s: no result by %s - outcome unknown", day, until.isoformat())
            return outcomes
        sleep(FOLLOW_POLL_SECONDS)


def _summary(outcomes: List[DayOutcome]) -> str:
    words = {"booked": "✅ booked", "covered": "✅ already booked", "would-book": "🟢 free (dry-run)",
             "no-slot": "nothing free",
             "failed": "⚠️ rejected/failed", "unconfirmed": "⚠️ unconfirmed", "error": "❌ error",
             "unknown": "❓ unknown"}
    return "; ".join(f"{_short(o.day)} {words.get(o.status, o.status)}" for o in outcomes)


def _released_days(cfg: Config, bookings: list, now: datetime, target: datetime, rehearsal: bool) -> List[str]:
    if not rehearsal:
        return booking.pending_days(cfg, bookings, now, released_at=target)
    # Rehearsal: next week's target days as if nothing were booked - display only, never booked.
    days = booking.pending_days(cfg, [], now, released_at=target)
    this_sunday = (now.date() + timedelta(days=6 - now.weekday())).isoformat()
    return [d for d in days if d > this_sunday] or days


def run_burst(cfg: Config, env: Dict[str, str], now: datetime, notifier: Notifier, state: State,
              now_fn: Callable[[], datetime],
              api_factory: Callable[[str, str], ParkingApi] = ParkingApi,
              spawn: Callable[..., "subprocess.Popen"] = spawn_worker,
              open_observers: Callable[..., Optional[Observers]] = Observers.open,
              sleep: Callable[[float], None] = time_module.sleep,
              run_dir: Optional[Path] = None, rehearsal: bool = False) -> RunResult:
    """Arm, run and report one burst (see the module docstring).

    Marks and saves the once-per-day guard first, so an overlapping tick stays
    out; the running marker then covers the burst until its report is sent.
    """
    burst = cfg.thursday_burst
    assert burst is not None
    if not rehearsal:
        state.mark_alert(BURST_STATE_KEY, now)
        save_state(STATE_PATH, state)

    token, stop = report.check_token(env, state, notifier, now)
    if stop is not None:
        return stop
    endpoint = env.get("PARKING_API_URL", "")
    plate = env.get("PARKING_LICENSE_PLATE", "")
    if not endpoint or not plate:
        raise SystemExit("PARKING_API_URL and PARKING_LICENSE_PLATE must be set in .env")

    target = (now + timedelta(seconds=REHEARSAL_LEAD_SECONDS)).replace(microsecond=0) if rehearsal \
        else burst_target(burst, now)
    deadline = target + timedelta(seconds=REHEARSAL_DURATION_SECONDS if rehearsal else burst.duration_seconds)
    try:
        # Read before the target, not at it: the 09-24 burst's first answer took 37 s at 16:00.
        bookings = api_factory(endpoint, token).my_bookings()
    except AuthError:
        logger.error("❌ burst: token rejected (401/403)")
        report.alert(state, notifier, "login-needed",
                     "Parking: the site rejected the token. Run: python -m parking.login", now)
        return RunResult("error", error="token rejected")
    except (RateLimited, ServerError, requests.RequestException, ApiError) as exc:
        logger.error("❌ burst: reading my bookings failed (%s); not arming", type(exc).__name__)
        return RunResult("error", error=type(exc).__name__)
    # All covered after this release -> no site calls until the next one.
    quiet_until = planner.next_release(target + timedelta(seconds=1), booking.release_time(cfg))
    days = _released_days(cfg, bookings, now, target, rehearsal)
    if not days:
        logger.info("ℹ️ burst: every day released at %s is already booked; nothing to do", target.isoformat())
        state.next_allowed = quiet_until
        return RunResult("nothing-pending")

    run_dir = run_dir or RUN_DIR
    run_dir.mkdir(parents=True, exist_ok=True)
    for stale in [*run_dir.glob("*.jsonl"), *run_dir.glob("*.log")]:
        stale.unlink(missing_ok=True)
    marker = run_dir / "running.json"
    marker.write_text(json.dumps({"started": now.isoformat(), "rehearsal": rehearsal}), encoding="utf-8")
    kind = "rehearsal" if rehearsal else ("Thursday" if now.weekday() == WEEKDAYS.index("thu") else "trial")
    logger.info("⚡ %s burst armed for %s: one worker per day %s", kind, target.isoformat(), days)
    try:
        # Workers first: Chrome starting or a slow Telegram send must never make one late for the target.
        workers = {day: spawn(day, target, deadline, cfg.dry_run, run_dir) for day in days}
        observers = open_observers(env.get("PARKING_URL", ""), days, PROFILE_DIR)
        notifier.send(f"⚡ Parking {kind} burst armed ({'dry-run, no bookings' if cfg.dry_run else 'LIVE, will book'}): "
                      f"one worker per day for {', '.join(_short(d) for d in days)}, checking every "
                      f"{burst.poll_interval_seconds:g}s from {target:%H:%M:%S}.", disposable=True)
        found = follow_workers(workers, run_dir, deadline + timedelta(seconds=WORKER_GRACE_SECONDS),
                               notifier, observers, sleep, now_fn)
        if observers is not None:
            sleep(OBSERVER_HOLD_SECONDS)
            observers.close()  # before the report: its screenshots need the same browser profile
        outcomes = [found[day] for day in days]
        logger.info("⚡ burst finished: %s", _summary(outcomes))
        notifier.ping(f"⚡ burst finished: {_summary(outcomes)}")
        report.report_outcomes(outcomes, cfg, state, notifier, now_fn(),
                               "" if rehearsal else env.get("PARKING_URL", ""))
    finally:
        # Only now: a tick during the report would fight it for the browser profile, state.json and the chat.
        marker.unlink(missing_ok=True)

    result = RunResult("checked", pending=days,
                       booked=[o.day for o in outcomes if o.status == "booked"],
                       would_book=[o.day for o in outcomes if o.status == "would-book"])
    broken = [o for o in outcomes if o.status in ("error", "unknown")]
    if broken:
        result.status, result.error = "error", "; ".join(f"{o.day}: {o.error}" for o in broken)
    elif all(o.status == "no-slot" for o in outcomes):
        result.status = "no-slots"
    if not rehearsal and all(o.status in ("booked", "covered") for o in outcomes):
        state.next_allowed = quiet_until
        logger.info("ℹ️ all covered: no site calls until the %s release", state.next_allowed.isoformat())
    return result


def rehearse(cfg: Config, env: Dict[str, str], now_fn: Callable[[], datetime]) -> RunResult:
    """The burst, starting in a few seconds, in forced dry-run: windows and workers
    run for real against the site's read-only queries, but `createBooking` is never
    called, state.json is untouched and nothing goes to Telegram."""
    now = now_fn()
    if is_running(now):
        logger.error("❌ a burst is running right now; not rehearsing on top of it")
        return RunResult("error", error="a burst is running")
    if cfg.thursday_burst is None:
        logger.error("❌ no thursday_burst in config.json; nothing to rehearse")
        return RunResult("error", error="no thursday_burst config")
    logger.info("ℹ️ rehearsal: dry-run forced, nothing will be booked, cancelled or sent")
    return run_burst(replace(cfg, dry_run=True), env, now, LogNotifier(), State(), now_fn=now_fn,
                     rehearsal=True)
