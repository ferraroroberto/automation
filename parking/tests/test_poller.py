import random
from datetime import datetime, timedelta, timezone
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

    def send(self, text):
        self.sent.append(text)
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
                "treat_placeless_as_covered": False, "jitter_max_seconds": 0})
    path = tmp_path / "config.json"
    path.write_text(json.dumps(raw))
    return load_config(path)


@pytest.fixture
def env():
    token = make_token(datetime.now(timezone.utc) + timedelta(days=20))
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
    env = dict(env, PARKING_TOKEN=make_token(datetime.now(timezone.utc) + timedelta(days=2)))
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


def test_outside_active_hours_does_nothing(cfg, env):
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
