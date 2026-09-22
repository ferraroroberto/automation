import random
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pytest

from parking import poller
from parking.api import ApiError, AuthError
from parking.config import load_config
from parking.state import State
from parking.tests.test_planner_auth import make_token

NOW = datetime(2026, 9, 20, 12, 0, tzinfo=ZoneInfo("Europe/Madrid"))  # Sunday; next Mon = 21st


class FakeNotifier:
    def __init__(self):
        self.sent = []
        self.sent_files = []
        self.sent_file_groups = []

    def send(self, text):
        self.sent.append(text)
        return True

    def send_file(self, text, path):
        self.sent_files.append((text, path))
        return True

    def send_files(self, text, paths):
        self.sent_file_groups.append((text, list(paths)))
        return True


class FakeApi:
    def __init__(self, bookings, slots, create_errors=None):
        self.bookings = bookings
        self.slots = slots  # (center_id, size_id) -> [raw days]
        self.create_errors = create_errors or {}
        self.created = []
        self.slot_calls = []

    def my_bookings(self):
        return list(self.bookings)

    def slot_days(self, center_id, size_id, type_):
        self.slot_calls.append((center_id, size_id))
        return self.slots.get((center_id, size_id), [])

    def create_booking(self, center_id, size_id, type_, plate, raw_day):
        if (center_id, size_id) in self.create_errors:
            raise self.create_errors[(center_id, size_id)]
        self.created.append((center_id, size_id, plate, raw_day))
        self.bookings.append({"day": raw_day, "state": "ACTIVE", "place": {"place": "1"}})
        return []


@pytest.fixture
def cfg(tmp_path):
    import json
    from pathlib import Path

    raw = json.loads((Path(__file__).resolve().parent.parent / "config.json").read_text())
    raw.update({"weekdays": ["mon"], "horizon_days": 7, "dry_run": False,
                "treat_placeless_as_covered": False, "jitter_max_seconds": 0,
                "sweep_report_until": None})
    path = tmp_path / "config.json"
    path.write_text(json.dumps(raw))
    return load_config(path)


@pytest.fixture
def env():
    token = make_token(NOW + timedelta(days=20))
    return {"PARKING_TOKEN": token, "PARKING_API_URL": "https://example.invalid/graphql",
            "PARKING_LICENSE_PLATE": "TEST123"}


def run(cfg, env, api, state=None, notifier=None, now=NOW):
    notifier = notifier or FakeNotifier()
    state = state or State()
    result = poller.run_once(cfg, env, now, notifier, state, api_factory=lambda *_: api,
                             sleep=lambda _s: None, rng=random.Random(0))
    return result, notifier, state


MON = "2026-09-21T10:00:00.000Z"


def test_dry_run_never_calls_create_booking(cfg, env):
    from dataclasses import replace

    api = FakeApi([], {(1, 3): [MON]})
    result, notifier, _ = run(replace(cfg, dry_run=True), env, api)
    assert api.created == []
    assert result.would_book == ["2026-09-21"]
    assert notifier.sent and notifier.sent[0].startswith("[dry-run]")


def test_books_one_day_per_call_using_first_preferred_combo(cfg, env):
    api = FakeApi([], {(1, 3): [MON], (2, 2): [MON]})
    result, notifier, _ = run(cfg, env, api)
    assert api.created == [(1, 3, "TEST123", MON)]
    assert result.booked == ["2026-09-21"]
    assert notifier.sent == ["Parking booked: 2026-09-21 at 22@ (Large)"]


def test_already_booked_day_is_never_booked_twice(cfg, env):
    booked = [{"day": MON, "state": "ACTIVE", "place": {"place": "12"}}]
    api = FakeApi(booked, {(1, 3): [MON]})
    result, _, _ = run(cfg, env, api)
    assert api.created == []
    assert api.slot_calls == []  # nothing pending -> no availability queries at all
    assert result.status == "nothing-pending"


def test_placeless_active_booking_blocks_a_second_booking_when_flag_on(cfg, env):
    from dataclasses import replace

    placeless = [{"day": MON, "state": "ACTIVE", "place": None}]
    api = FakeApi(placeless, {(1, 3): [MON]})
    run(replace(cfg, treat_placeless_as_covered=True), env, api)
    assert api.created == []


def test_falls_through_to_next_slot_when_booking_rejected(cfg, env):
    api = FakeApi([], {(1, 3): [MON], (2, 2): [MON]}, create_errors={(1, 3): ApiError("taken")})
    result, _, _ = run(cfg, env, api)
    assert api.created == [(2, 2, "TEST123", MON)]
    assert result.booked == ["2026-09-21"]


def test_no_slots_is_silent(cfg, env):
    api = FakeApi([], {})
    result, notifier, _ = run(cfg, env, api)
    assert result.status == "no-slots" and notifier.sent == []
    assert len(api.slot_calls) == 6  # 2 centers x 3 sizes


def test_missing_token_alerts_once_per_day(cfg):
    state = State()
    notifier = FakeNotifier()
    for _ in range(2):
        run(cfg, {}, FakeApi([], {}), state=state, notifier=notifier)
    assert len(notifier.sent) == 1 and "login needed" in notifier.sent[0]


def test_expiring_token_alerts_but_still_polls(cfg, env):
    env = dict(env, PARKING_TOKEN=make_token(NOW + timedelta(days=2)))
    api = FakeApi([], {})
    result, notifier, _ = run(cfg, env, api)
    assert result.status == "no-slots"
    assert any("expires in" in text for text in notifier.sent)


def test_auth_error_alerts_and_backs_off(cfg, env):
    class Rejecting(FakeApi):
        def my_bookings(self):
            raise AuthError("HTTP 401")

    result, notifier, state = run(cfg, env, Rejecting([], {}))
    assert result.status == "error" and state.next_allowed is not None
    assert any("rejected the token" in text for text in notifier.sent)
    again, _, _ = run(cfg, env, Rejecting([], {}), state=state, notifier=notifier)
    assert again.status == "backoff"


def test_off_window_does_nothing(cfg, env):
    api = FakeApi([], {(1, 3): [MON]})
    late = NOW.replace(hour=23)
    result, _, _ = run(cfg, env, api, now=late)
    assert result.status == "inactive" and api.slot_calls == []


def test_repeated_errors_alert_at_threshold(cfg, env):
    class Failing(FakeApi):
        def my_bookings(self):
            raise ApiError("boom")

    state = State()
    notifier = FakeNotifier()
    now = NOW
    for _ in range(cfg.repeated_error_threshold):
        run(cfg, env, Failing([], {}), state=state, notifier=notifier, now=now)
        now = state.next_allowed + timedelta(minutes=1)
    assert sum("consecutive errors" in t for t in notifier.sent) == 1


def test_token_value_never_logged(cfg, env, caplog):
    caplog.set_level("DEBUG")
    run(cfg, env, FakeApi([], {}))
    assert env["PARKING_TOKEN"] not in caplog.text


def test_all_booking_attempts_failing_alerts_once_per_hour(cfg, env):
    errors = {(1, 3): ApiError("bad date format"), (2, 2): ApiError("bad date format")}
    state = State()
    notifier = FakeNotifier()
    first = run(cfg, env, FakeApi([], {(1, 3): [MON], (2, 2): [MON]}, errors),
                state=state, notifier=notifier)[0]
    assert first.booked == []
    run(cfg, env, FakeApi([], {(1, 3): [MON], (2, 2): [MON]}, errors),
        state=state, notifier=notifier, now=NOW + timedelta(minutes=15))
    failed = [t for t in notifier.sent if "booking failed" in t]
    assert len(failed) == 1 and "2026-09-21" in failed[0] and "bad date format" in failed[0]
    run(cfg, env, FakeApi([], {(1, 3): [MON]}, {(1, 3): ApiError("still bad")}),
        state=state, notifier=notifier, now=NOW + timedelta(hours=2))
    assert len([t for t in notifier.sent if "booking failed" in t]) == 2


def test_no_failure_alert_when_a_later_slot_books(cfg, env):
    api = FakeApi([], {(1, 3): [MON], (2, 2): [MON]}, {(1, 3): ApiError("taken")})
    _, notifier, _ = run(cfg, env, api)
    assert not any("booking failed" in t for t in notifier.sent)


def test_holiday_target_day_is_never_queried_or_booked(cfg, env):
    from dataclasses import replace

    api = FakeApi([], {(1, 3): [MON]})
    result, notifier, _ = run(replace(cfg, skip_dates=["2026-09-21"]), env, api)
    assert result.status == "nothing-pending"
    assert api.slot_calls == [] and api.created == [] and notifier.sent == []


def at(hour, minute=0):
    return NOW.replace(hour=hour, minute=minute)


@pytest.mark.parametrize("hour, minute, expected", [
    (19, 0, 5), (22, 55, 5), (5, 0, 5), (8, 55, 5),  # evening and morning windows, start inclusive
    (23, 0, 0), (2, 0, 0), (4, 59, 0),                # overnight off window wraps midnight, end exclusive
    (9, 0, 15), (12, 0, 15), (18, 55, 15),            # default elsewhere
])
def test_poll_interval_by_time_window(cfg, hour, minute, expected):
    assert poller.poll_interval(cfg, at(hour, minute)) == expected


def test_success_schedules_the_next_run_by_window(cfg, env):
    _, _, state = run(cfg, env, FakeApi([], {}), now=at(21))
    assert state.next_allowed == at(21) + timedelta(minutes=5, seconds=-cfg.jitter_max_seconds)
    _, _, state = run(cfg, env, FakeApi([], {}), now=at(12))
    assert state.next_allowed == at(12) + timedelta(minutes=15, seconds=-cfg.jitter_max_seconds)


def test_run_arriving_before_it_is_due_makes_no_api_call(cfg, env):
    api = FakeApi([], {})
    _, _, state = run(cfg, env, api, now=at(12))
    api.slot_calls.clear()
    early, _, _ = run(cfg, env, api, state=state, now=at(12, 5))
    assert early.status == "backoff" and api.slot_calls == []
    due, _, _ = run(cfg, env, api, state=state, now=at(12, 15))
    assert due.status == "no-slots"


THU = datetime(2026, 9, 24, 16, 2, tzinfo=ZoneInfo("Europe/Madrid"))  # Thursday
THU_MON = "2026-09-28T10:00:00.000Z"  # 24 Sept is a holiday; the next Monday is not


def test_thursday_16h_window_sends_a_live_log(cfg, env):
    from dataclasses import replace

    api = FakeApi([], {(1, 3): [THU_MON]})
    weekdays_cfg = replace(cfg, weekdays=["mon"])
    result, notifier, _ = run(weekdays_cfg, env, api, now=THU)
    assert result.booked == ["2026-09-28"]
    log = " | ".join(notifier.sent)
    for step in ("poll started", "Checked my bookings", "Looking for free slots", "Free slot found",
                 "Booking 2026-09-28", "Parking booked", "Poll finished"):
        assert step in log
    assert notifier.sent.index(next(t for t in notifier.sent if "Booking" in t)) < notifier.sent.index(
        "Parking booked: 2026-09-28 at 22@ (Large)")


def test_quiet_polls_only_notify_on_booking(cfg, env):
    from dataclasses import replace

    api = FakeApi([], {(1, 3): [THU_MON]})
    result, notifier, _ = run(replace(cfg, weekdays=["mon"]), env, api, now=THU.replace(hour=12))
    assert result.booked == ["2026-09-28"]
    assert notifier.sent == ["Parking booked: 2026-09-28 at 22@ (Large)"]


def test_thursday_window_is_thursday_only(cfg):
    monday = THU + timedelta(days=4)
    assert poller.is_verbose(cfg, THU) and not poller.is_verbose(cfg, monday)
    assert poller.poll_interval(cfg, THU.replace(hour=15, minute=50)) == 5
    assert poller.poll_interval(cfg, monday.replace(hour=15, minute=50)) == 15


def test_run_just_before_16h_does_not_block_the_16h_poll(cfg, env):
    from dataclasses import replace

    quiet = replace(cfg, jitter_max_seconds=120)
    _, _, state = run(quiet, env, FakeApi([], {}), now=THU.replace(hour=15, minute=58))
    assert not state.in_backoff(THU)  # 16:02; a 15-minute interval would block until 16:11


def test_sweep_report_sends_one_line_per_sweep_until_the_end_time(cfg, env):
    from dataclasses import replace

    until = NOW + timedelta(hours=1)
    trial = replace(cfg, sweep_report_until=until)
    _, notifier, _ = run(trial, env, FakeApi([], {}))
    assert notifier.sent == ["Parking sweep Sun 12:00 ✅ 1 open day(s), no free slot for them."]
    _, notifier, _ = run(trial, env, FakeApi([], {(1, 3): [MON]}))
    assert notifier.sent == ["Parking booked: 2026-09-21 at 22@ (Large)",
                             "Parking sweep Sun 12:00 ✅ 1 open day(s), booked: 2026-09-21."]
    _, notifier, _ = run(trial, env, FakeApi([], {}), now=until)
    assert notifier.sent == []


def test_sweep_report_says_when_a_sweep_fails_and_skips_early_runs(cfg, env):
    from dataclasses import replace

    from parking.api import ServerError

    class Broken(FakeApi):
        def my_bookings(self):
            raise ServerError("503")

    trial = replace(cfg, sweep_report_until=NOW + timedelta(hours=1))
    _, notifier, state = run(trial, env, Broken([], {}))
    assert notifier.sent == ["Parking sweep Sun 12:00 ❌ failed (ServerError), will retry."]
    _, notifier, _ = run(trial, env, FakeApi([], {}), state=state)
    assert notifier.sent == []  # still backing off: no sweep, no report


def _make_pngs(tmp_path, n):
    paths = []
    for i in range(n):
        path = tmp_path / f"shot{i}.png"
        path.write_bytes(b"\x89PNG")
        paths.append(path)
    return paths


def test_booked_notification_attaches_all_screenshots_as_one_message(cfg, env, monkeypatch, tmp_path):
    shots = _make_pngs(tmp_path, 4)
    monkeypatch.setattr(poller.site_screenshot, "capture_all", lambda *a, **k: list(shots))
    api = FakeApi([], {(1, 3): [MON]})
    notifier = FakeNotifier()
    run(cfg, dict(env, PARKING_URL="https://example.invalid/book"), api, notifier=notifier)
    assert notifier.sent == [] and notifier.sent_files == []
    assert notifier.sent_file_groups == [
        ("Parking booked: 2026-09-21 at 22@ (Large)", [str(p) for p in shots])]
    assert not any(p.exists() for p in shots)  # deleted after the send attempt


def test_booked_notification_uses_single_file_send_for_one_screenshot(cfg, env, monkeypatch, tmp_path):
    shots = _make_pngs(tmp_path, 1)
    monkeypatch.setattr(poller.site_screenshot, "capture_all", lambda *a, **k: list(shots))
    api = FakeApi([], {(1, 3): [MON]})
    notifier = FakeNotifier()
    run(cfg, dict(env, PARKING_URL="https://example.invalid/book"), api, notifier=notifier)
    assert notifier.sent_file_groups == []
    assert notifier.sent_files == [("Parking booked: 2026-09-21 at 22@ (Large)", str(shots[0]))]
    assert not shots[0].exists()


def test_booked_notification_falls_back_to_text_when_nothing_captured(cfg, env, monkeypatch):
    monkeypatch.setattr(poller.site_screenshot, "capture_all", lambda *a, **k: [])
    api = FakeApi([], {(1, 3): [MON]})
    notifier = FakeNotifier()
    run(cfg, dict(env, PARKING_URL="https://example.invalid/book"), api, notifier=notifier)
    assert notifier.sent == ["Parking booked: 2026-09-21 at 22@ (Large)"]
    assert notifier.sent_files == [] and notifier.sent_file_groups == []


def test_booked_notification_falls_back_to_text_when_group_send_fails(cfg, env, monkeypatch, tmp_path):
    shots = _make_pngs(tmp_path, 4)
    monkeypatch.setattr(poller.site_screenshot, "capture_all", lambda *a, **k: list(shots))

    class FailingGroupNotifier(FakeNotifier):
        def send_files(self, text, paths):
            self.sent_file_groups.append((text, list(paths)))
            return False

    api = FakeApi([], {(1, 3): [MON]})
    notifier = FailingGroupNotifier()
    run(cfg, dict(env, PARKING_URL="https://example.invalid/book"), api, notifier=notifier)
    assert notifier.sent == ["Parking booked: 2026-09-21 at 22@ (Large)"]
    assert len(notifier.sent_file_groups) == 1
    assert not any(p.exists() for p in shots)  # cleaned up even though the send failed


def test_no_screenshot_attempt_without_parking_url(cfg, env, monkeypatch):
    called = []
    monkeypatch.setattr(poller.site_screenshot, "capture_all",
                        lambda *a, **k: called.append(1) or [])
    api = FakeApi([], {(1, 3): [MON]})
    notifier = FakeNotifier()
    run(cfg, env, api, notifier=notifier)  # env fixture has no PARKING_URL
    assert called == []
    assert notifier.sent == ["Parking booked: 2026-09-21 at 22@ (Large)"]
    assert notifier.sent_files == [] and notifier.sent_file_groups == []


def test_booked_notification_passes_both_configured_centers_to_capture(cfg, env, monkeypatch):
    seen = []
    monkeypatch.setattr(poller.site_screenshot, "capture_all",
                        lambda centers, *a, **k: seen.append(list(centers)) or [])
    api = FakeApi([], {(1, 3): [MON]})
    run(cfg, dict(env, PARKING_URL="https://example.invalid/book"), api)
    assert seen == [["22@", "CINC"]]


def test_booking_failed_alert_attaches_screenshots_when_available(cfg, env, monkeypatch, tmp_path):
    shots = _make_pngs(tmp_path, 4)
    monkeypatch.setattr(poller.site_screenshot, "capture_all", lambda *a, **k: list(shots))
    errors = {(1, 3): ApiError("bad date format"), (2, 2): ApiError("bad date format")}
    notifier = FakeNotifier()
    run(cfg, dict(env, PARKING_URL="https://example.invalid/book"),
        FakeApi([], {(1, 3): [MON], (2, 2): [MON]}, errors), notifier=notifier)
    assert notifier.sent == []
    assert len(notifier.sent_file_groups) == 1 and "booking failed" in notifier.sent_file_groups[0][0]


class FakeClock:
    """A controllable clock: `sleep()` advances `now` instead of blocking."""

    def __init__(self, start):
        self.now = start

    def sleep(self, seconds):
        self.now += timedelta(seconds=seconds)

    def now_fn(self):
        return self.now


THU_1602 = datetime(2026, 9, 24, 16, 2, tzinfo=ZoneInfo("Europe/Madrid"))
BURST_TARGET = THU_1602.replace(hour=16, minute=0, second=0, microsecond=0)


def test_burst_watch_band_covers_the_lead_in_and_the_burst_itself(cfg):
    burst = cfg.thursday_burst
    assert burst.watch_before_minutes == 6 and burst.duration_seconds == 90
    assert poller._in_burst_watch_band(burst, BURST_TARGET - timedelta(minutes=6)) is True
    assert poller._in_burst_watch_band(burst, BURST_TARGET) is True
    assert poller._in_burst_watch_band(burst, BURST_TARGET + timedelta(seconds=89)) is True
    assert poller._in_burst_watch_band(burst, BURST_TARGET - timedelta(minutes=6, seconds=1)) is False
    assert poller._in_burst_watch_band(burst, BURST_TARGET + timedelta(seconds=91)) is False


def test_burst_watch_band_is_thursday_only(cfg):
    burst = cfg.thursday_burst
    not_thursday = BURST_TARGET - timedelta(days=1)
    assert poller._in_burst_watch_band(burst, not_thursday) is False


def test_sleep_until_precise_busy_polls_the_final_stretch():
    clock = FakeClock(BURST_TARGET - timedelta(seconds=10))
    calls = []
    def sleep(seconds):
        calls.append(seconds)
        clock.sleep(seconds)
    poller._sleep_until_precise(BURST_TARGET, sleep, clock.now_fn)
    assert clock.now == BURST_TARGET
    assert calls[0] == pytest.approx(10 - poller.BURST_FINE_APPROACH_SECONDS)
    assert all(c == poller.BURST_FINE_POLL_SECONDS for c in calls[1:])
    assert len(calls) > 1  # actually busy-polled, not one long sleep


def test_sleep_until_precise_skips_the_coarse_sleep_when_already_close():
    clock = FakeClock(BURST_TARGET - timedelta(seconds=1))
    calls = []
    poller._sleep_until_precise(BURST_TARGET, lambda s: (calls.append(s), clock.sleep(s)), clock.now_fn)
    assert all(c == poller.BURST_FINE_POLL_SECONDS for c in calls)


def test_burst_marks_and_saves_state_before_sleeping(cfg, env, tmp_path, monkeypatch):
    monkeypatch.setattr(poller, "STATE_PATH", tmp_path / "state.json")
    clock = FakeClock(BURST_TARGET - timedelta(seconds=5))
    state = State()
    api = FakeApi([], {})
    poller._run_burst(cfg, env, clock.now, FakeNotifier(), state, lambda *_: api,
                      clock.sleep, random.Random(0), clock.now_fn)
    assert not state.alert_due(poller.BURST_STATE_KEY, clock.now, poller.BURST_COOLDOWN)
    saved = poller.load_state(tmp_path / "state.json")
    assert poller.BURST_STATE_KEY in saved.alerts  # persisted before the sleep, not just in memory


def test_burst_polls_repeatedly_until_booked_then_stops_early(cfg, env, tmp_path, monkeypatch):
    monkeypatch.setattr(poller, "STATE_PATH", tmp_path / "state.json")
    clock = FakeClock(BURST_TARGET - timedelta(seconds=3))
    state = State()
    notifier = FakeNotifier()

    class FlakyApi(FakeApi):
        """No slot for the first two cycles, then one appears - proves the
        burst actually re-checks the site instead of stopping after one look."""

        def __init__(self):
            super().__init__([], {})
            self.cycle = 0

        def my_bookings(self):
            self.cycle += 1
            return super().my_bookings()

        def slot_days(self, center_id, size_id, type_):
            self.slot_calls.append((center_id, size_id))
            if self.cycle < 3:
                return []
            # 2026-09-24 is a Thursday and a configured skip_date; the next
            # Monday (weekdays=["mon"]) within the horizon is 2026-09-28.
            return {(1, 3): ["2026-09-28T10:00:00.000Z"]}.get((center_id, size_id), [])

    api = FlakyApi()
    result = poller._run_burst(cfg, env, clock.now, notifier, state, lambda *_a: api,
                               clock.sleep, random.Random(0), clock.now_fn)
    assert result.booked == ["2026-09-28"]
    assert api.cycle >= 3  # re-checked at least until the slot appeared
    assert any("Parking booked" in t for t in notifier.sent + [g[0] for g in notifier.sent_file_groups]
              + [f[0] for f in notifier.sent_files])


def test_burst_stops_on_the_first_exception_without_retrying(cfg, env, tmp_path, monkeypatch):
    monkeypatch.setattr(poller, "STATE_PATH", tmp_path / "state.json")
    clock = FakeClock(BURST_TARGET - timedelta(seconds=1))
    state = State()

    class BrokenApi(FakeApi):
        def __init__(self):
            super().__init__([], {})
            self.attempts = 0

        def my_bookings(self):
            self.attempts += 1
            raise ApiError("boom")

    api = BrokenApi()
    result = poller._run_burst(cfg, env, clock.now, FakeNotifier(), state, lambda *_a: api,
                               clock.sleep, random.Random(0), clock.now_fn)
    assert result.status == "error"
    assert api.attempts == 1  # one attempt, no retry into the same failure


def test_burst_respects_the_duration_cap_when_nothing_ever_frees_up(cfg, env, tmp_path, monkeypatch):
    monkeypatch.setattr(poller, "STATE_PATH", tmp_path / "state.json")
    clock = FakeClock(BURST_TARGET)
    state = State()

    class CountingApi(FakeApi):
        def __init__(self):
            super().__init__([], {})  # never any free slot
            self.cycles = 0

        def my_bookings(self):
            self.cycles += 1
            return super().my_bookings()

    api = CountingApi()
    result = poller._run_burst(cfg, env, clock.now, FakeNotifier(), state, lambda *_a: api,
                               clock.sleep, random.Random(0), clock.now_fn)
    assert result.status == "no-slots"
    expected_iterations = cfg.thursday_burst.duration_seconds / cfg.thursday_burst.poll_interval_seconds
    assert api.cycles <= expected_iterations + 1  # bounded, didn't loop forever
    assert clock.now >= BURST_TARGET + timedelta(seconds=cfg.thursday_burst.duration_seconds)


def test_burst_not_armed_without_a_token(cfg, tmp_path, monkeypatch):
    monkeypatch.setattr(poller, "STATE_PATH", tmp_path / "state.json")
    clock = FakeClock(BURST_TARGET)
    state = State()
    result = poller._run_burst(cfg, {}, clock.now, FakeNotifier(), state, lambda *_: FakeApi([], {}),
                               clock.sleep, random.Random(0), clock.now_fn)
    assert result.status == "no-token"


def test_sweep_report_attaches_screenshots_when_available(cfg, env, monkeypatch, tmp_path):
    from dataclasses import replace

    shots = _make_pngs(tmp_path, 4)
    monkeypatch.setattr(poller.site_screenshot, "capture_all", lambda *a, **k: list(shots))
    trial = replace(cfg, sweep_report_until=NOW + timedelta(hours=1))
    notifier = FakeNotifier()
    run(trial, dict(env, PARKING_URL="https://example.invalid/book"), FakeApi([], {}), notifier=notifier)
    assert notifier.sent == []
    assert len(notifier.sent_file_groups) == 1 and "Parking sweep" in notifier.sent_file_groups[0][0]
