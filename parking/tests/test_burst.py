import json
import random
from dataclasses import replace
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pytest

from parking import burst, burst_worker
from parking.booking import DayOutcome
from parking.state import State, load_state
from parking.tests.test_poller import FakeApi, FakeNotifier, cfg, env  # noqa: F401 (fixtures)

TZ = ZoneInfo("Europe/Madrid")
TARGET = datetime(2026, 9, 24, 16, 0, tzinfo=TZ)  # a Thursday release (the 24th itself is a holiday)
MON = "2026-09-28T10:00:00.000Z"
THU = "2026-10-01T10:00:00.000Z"


class FakeClock:
    """A controllable clock: `sleep()` advances `now` instead of blocking."""

    def __init__(self, start):
        self.now = start

    def sleep(self, seconds):
        self.now += timedelta(seconds=seconds)

    def now_fn(self):
        return self.now


# --- when to burst ------------------------------------------------------------------------------

def test_watch_band_covers_the_lead_in_and_the_burst_itself(cfg):
    band = cfg.thursday_burst
    assert band.watch_before_minutes == 6 and band.duration_seconds == 90
    assert burst.in_watch_band(band, TARGET - timedelta(minutes=6)) is True
    assert burst.in_watch_band(band, TARGET + timedelta(seconds=89)) is True
    assert burst.in_watch_band(band, TARGET - timedelta(minutes=6, seconds=1)) is False
    assert burst.in_watch_band(band, TARGET + timedelta(seconds=91)) is False


def test_watch_band_is_thursday_or_a_trial_date(cfg):
    wednesday = TARGET - timedelta(days=1)
    assert burst.in_watch_band(replace(cfg.thursday_burst, trial_dates=()), wednesday) is False
    trial = replace(cfg.thursday_burst, trial_dates=(wednesday.date(),))
    assert burst.in_watch_band(trial, wednesday) is True
    assert burst.in_watch_band(trial, TARGET) is True


def test_a_tick_during_a_burst_is_skipped(cfg, tmp_path):
    state = State()
    armed_at = TARGET - timedelta(minutes=4)
    assert burst.burst_action(cfg, state, armed_at, tmp_path) == "burst"
    state.mark_alert(burst.BURST_STATE_KEY, armed_at)
    assert burst.burst_action(cfg, state, TARGET + timedelta(seconds=30), tmp_path) == "skip"
    assert burst.burst_action(cfg, state, TARGET + timedelta(minutes=5), tmp_path) == "normal"


def test_the_running_marker_keeps_ticks_out_past_the_band(cfg, tmp_path):
    """Workers and the report outlast the watch band: a tick must not poll alongside them."""
    (tmp_path / "running.json").write_text(json.dumps({"started": TARGET.isoformat()}))
    later = TARGET + timedelta(minutes=5)
    assert burst.burst_action(cfg, State(), later, tmp_path) == "skip"
    assert burst.burst_action(cfg, State(), TARGET + burst.RUNNING_MAX_AGE, tmp_path) == "normal"  # stale


def test_sleep_until_precise_busy_polls_the_final_stretch():
    clock = FakeClock(TARGET - timedelta(seconds=10))
    calls = []

    def sleep(seconds):
        calls.append(seconds)
        clock.sleep(seconds)

    burst.sleep_until_precise(TARGET, sleep, clock.now_fn)
    assert clock.now == TARGET
    assert calls[0] == pytest.approx(10 - burst.FINE_APPROACH_SECONDS)
    assert len(calls) > 1 and all(c == burst.FINE_POLL_SECONDS for c in calls[1:])


# --- one day's worker ---------------------------------------------------------------------------

def work(cfg, api, day="2026-09-28", dry_run=False, start=TARGET - timedelta(seconds=30)):
    clock = FakeClock(start)
    lines = []
    outcome = burst_worker.work_day(api, cfg, day, "TEST123", TARGET, TARGET + timedelta(seconds=90),
                                    dry_run, lines.append, clock.sleep, clock.now_fn, random.Random(0))
    return outcome, lines, clock


def test_worker_waits_for_the_target_then_books_the_first_free_slot(cfg):
    api = FakeApi([], {(1, 3): [MON], (2, 2): [MON]})
    outcome, lines, _ = work(cfg, api)
    assert outcome == DayOutcome("2026-09-28", "booked", "2026-09-28 at 22@ (Large)")
    assert api.created == [(1, 3, "TEST123", MON)]
    assert lines == ["#1 16:00:00.00 ✅ booked 2026-09-28 at 22@ (Large)"]


def test_worker_moves_to_the_next_slot_when_one_is_rejected(cfg):
    api = FakeApi([], {(1, 3): [MON], (2, 2): [MON]}, create_states={(1, 3): "REJECTED"})
    outcome, _, _ = work(cfg, api)
    assert outcome.status == "booked" and outcome.label == "2026-09-28 at CINC (Medium)"
    assert api.created == [(1, 3, "TEST123", MON), (2, 2, "TEST123", MON)]


def test_worker_never_retries_a_refused_slot(cfg):
    """Each REJECTED attempt leaves a record on the site: one per slot, not one per check."""
    api = FakeApi([], {(1, 3): [MON]}, create_states={(1, 3): "REJECTED"})
    outcome, lines, _ = work(cfg, api)
    assert api.created == [(1, 3, "TEST123", MON)]
    assert outcome.status == "no-slot" and len(lines) > 2  # kept checking the other slots


def test_worker_keeps_checking_until_the_deadline(cfg):
    api = FakeApi([], {})
    outcome, lines, clock = work(cfg, api)
    assert outcome.status == "no-slot" and api.created == []
    assert len(lines) > 10 and all("nothing free" in line for line in lines)
    assert clock.now <= TARGET + timedelta(seconds=90)


def test_dry_run_worker_never_books(cfg):
    api = FakeApi([], {(1, 3): [MON]})
    outcome, lines, _ = work(cfg, api, dry_run=True)
    assert outcome.status == "would-book" and api.created == []
    assert "free (dry-run)" in lines[0]


def test_worker_never_books_a_day_that_already_has_an_active_booking(cfg):
    booked = [{"day": MON, "state": "ACTIVE", "place": {"place": "7"}}]
    api = FakeApi(booked, {(1, 3): [MON]})
    outcome, lines, _ = work(cfg, api)
    assert outcome.status == "covered" and api.created == [] and api.slot_calls == [] and lines == []


def test_dry_run_worker_still_scans_a_booked_day(cfg):
    """The rehearsal shows the workers on days that are already booked, still without booking."""
    booked = [{"day": MON, "state": "ACTIVE", "place": {"place": "7"}}]
    api = FakeApi(booked, {(1, 3): [MON]})
    outcome, _, _ = work(cfg, api, dry_run=True)
    assert outcome.status == "would-book" and api.created == [] and api.slot_calls


# --- the orchestrator ---------------------------------------------------------------------------

class FakeProc:
    def __init__(self, code=0):
        self.returncode = code

    def poll(self):
        return self.returncode


class FakeObservers:
    def __init__(self):
        self.refreshed, self.closed = [], False

    def refresh(self, day):
        self.refreshed.append(day)

    def close(self):
        self.closed = True


def inline_spawn(api, spawned):
    """Run each worker in-process, each on its own clock, writing the real event file."""
    def spawn(day, target, deadline, dry_run, run_dir):
        spawned.append((day, dry_run))
        clock = FakeClock(target - timedelta(seconds=5))
        path = run_dir / f"{day}.jsonl"

        def write(record):
            with path.open("a", encoding="utf-8") as handle:
                handle.write(json.dumps(record) + "\n")

        outcome = burst_worker.work_day(api, burst_cfg[0], day, "TEST123", target, deadline, dry_run,
                                        lambda text: write({"text": text}), clock.sleep, clock.now_fn,
                                        random.Random(0))
        write({"done": True, "status": outcome.status, "label": outcome.label, "error": outcome.error})
        return FakeProc()
    return spawn


burst_cfg = [None]  # the config inline workers use (they load their own in real life)


def orchestrate(cfg, env, api, tmp_path, monkeypatch, spawn=None, rehearsal=False, now=None):
    monkeypatch.setattr(burst, "STATE_PATH", tmp_path / "state.json")
    burst_cfg[0] = cfg
    spawned, observers = [], FakeObservers()
    clock = FakeClock(now or TARGET - timedelta(minutes=3))
    notifier = FakeNotifier()
    state = State()
    result = burst.run_burst(cfg, env, clock.now, notifier, state, clock.now_fn, api_factory=lambda *_: api,
                             spawn=spawn or inline_spawn(api, spawned),
                             open_observers=lambda *_a: observers, sleep=clock.sleep,
                             run_dir=tmp_path / "burst", rehearsal=rehearsal)
    return result, spawned, observers, notifier, state


def test_burst_runs_one_worker_per_released_day_and_reports_after(cfg, env, tmp_path, monkeypatch):
    two_days = replace(cfg, weekdays=["mon", "thu"])  # Thu 24th is a holiday: Mon 28th + Thu 1st
    api = FakeApi([], {(1, 3): [MON], (2, 2): [THU]})
    result, spawned, observers, notifier, state = orchestrate(two_days, env, api, tmp_path, monkeypatch)
    assert spawned == [("2026-09-28", False), ("2026-10-01", False)]
    assert result.status == "checked" and result.booked == ["2026-09-28", "2026-10-01"]
    assert notifier.sent[0].startswith("⚡ Parking Thursday burst armed (LIVE, will book): one worker per day")
    assert "Parking booked: 2026-09-28 at 22@ (Large); 2026-10-01 at CINC (Medium)" in notifier.sent
    assert any(p.startswith("⚡ 28/09 #1 16:00:00") for p in notifier.pings)
    assert notifier.pings[-1] == "⚡ burst finished: 28/09 ✅ booked; 01/10 ✅ booked"
    assert observers.refreshed == ["2026-09-28", "2026-10-01"] and observers.closed
    assert not (tmp_path / "burst" / "running.json").exists()
    assert state.next_allowed == TARGET + timedelta(days=7)  # all covered: quiet until next week
    assert burst.BURST_STATE_KEY in load_state(tmp_path / "state.json").alerts


def test_burst_marks_and_saves_its_guard_before_anything_else(cfg, env, tmp_path, monkeypatch):
    def spawn(*_a):
        assert burst.BURST_STATE_KEY in load_state(tmp_path / "state.json").alerts
        return FakeProc(1)

    orchestrate(cfg, env, FakeApi([], {}), tmp_path, monkeypatch, spawn=spawn)


def test_burst_leaves_out_days_not_released_at_the_target(cfg, env, tmp_path, monkeypatch):
    wide = replace(cfg, horizon_days=30)
    result, spawned, *_ = orchestrate(wide, env, FakeApi([], {}), tmp_path, monkeypatch)
    assert [d for d, _ in spawned] == ["2026-09-28"]  # Oct 5th onwards opens next Thursday
    assert result.status == "no-slots"


def test_burst_with_every_day_booked_starts_no_worker(cfg, env, tmp_path, monkeypatch):
    booked = [{"day": MON, "state": "ACTIVE", "place": {"place": "7"}}]
    result, spawned, _, notifier, state = orchestrate(cfg, env, FakeApi(booked, {}), tmp_path, monkeypatch)
    assert result.status == "nothing-pending" and spawned == [] and notifier.sent == []
    assert state.next_allowed == TARGET + timedelta(days=7)


def test_a_worker_that_dies_without_a_result_is_an_error_not_a_success(cfg, env, tmp_path, monkeypatch):
    result, *_rest, notifier, _ = orchestrate(cfg, env, FakeApi([], {}), tmp_path, monkeypatch,
                                              spawn=lambda *_a: FakeProc(1))
    assert result.status == "error" and "without a result" in result.error
    assert any("ended error" in t for t in notifier.sent)


def test_a_silent_worker_is_unknown_after_the_grace(cfg, env, tmp_path, monkeypatch):
    result, *_rest, notifier, _ = orchestrate(cfg, env, FakeApi([], {}), tmp_path, monkeypatch,
                                              spawn=lambda *_a: FakeProc(None))
    assert result.status == "error" and "no result" in result.error
    assert any("ended unknown" in t for t in notifier.sent)


def test_burst_without_a_token_starts_nothing(cfg, tmp_path, monkeypatch):
    result, spawned, *_ = orchestrate(cfg, {}, FakeApi([], {}), tmp_path, monkeypatch)
    assert result.status == "no-token" and spawned == []


def test_rehearsal_never_books_and_writes_no_state(cfg, env, tmp_path, monkeypatch):
    """Days already booked still get a worker (display only), in dry-run, with no state touched."""
    booked = [{"day": MON, "state": "ACTIVE", "place": {"place": "7"}}]
    api = FakeApi(booked, {(1, 3): [MON]})
    dry = replace(cfg, dry_run=True)
    result, spawned, *_rest, state = orchestrate(dry, env, api, tmp_path, monkeypatch, rehearsal=True,
                                                 now=TARGET + timedelta(hours=1))
    assert spawned == [("2026-09-28", True)] and api.created == []
    assert result.would_book == ["2026-09-28"]
    assert not (tmp_path / "state.json").exists() and state.next_allowed is None  # its State is thrown away


def test_rehearse_forces_dry_run_and_stays_out_of_the_chat(cfg, env, monkeypatch, tmp_path):
    from parking.notify import LogNotifier

    seen = {}
    monkeypatch.setattr(burst, "RUN_DIR", tmp_path)
    monkeypatch.setattr(burst, "run_burst",
                        lambda c, e, n, notifier, s, now_fn, rehearsal: seen.update(
                            dry_run=c.dry_run, notifier=notifier, rehearsal=rehearsal) or "ran")
    live = replace(cfg, dry_run=False)
    assert burst.rehearse(live, env, lambda: TARGET) == "ran"
    assert seen["dry_run"] is True and seen["rehearsal"] is True
    assert isinstance(seen["notifier"], LogNotifier)
