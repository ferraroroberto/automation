from datetime import date, datetime, time, timedelta, timezone
from zoneinfo import ZoneInfo

from parking import auth, planner
from parking.planner import Combo


def make_token(exp: datetime) -> str:
    import base64
    import json

    def enc(obj: dict) -> str:
        return base64.urlsafe_b64encode(json.dumps(obj).encode()).decode().rstrip("=")

    return f"{enc({'alg': 'none'})}.{enc({'exp': int(exp.timestamp())})}.sig"


def test_target_dates_filters_weekdays_within_horizon():
    # 2026-09-21 is a Monday
    days = planner.target_dates(date(2026, 9, 21), ["mon", "thu", "fri"], 7)
    assert days == ["2026-09-21", "2026-09-24", "2026-09-25", "2026-09-28"]


def test_target_dates_accepts_long_names_and_case():
    assert planner.target_dates(date(2026, 9, 21), ["Monday"], 0) == ["2026-09-21"]


def test_covered_requires_place_by_default_flag_off():
    bookings = [
        {"day": "2026-09-21T10:00:00.000Z", "state": "ACTIVE", "place": {"place": None}},
        {"day": "2026-09-24T10:00:00.000Z", "state": "ACTIVE", "place": {"place": "49"}},
        {"day": "2026-09-25T10:00:00.000Z", "state": "CANCELLED", "place": {"place": "7"}},
    ]
    covered, placeless = planner.covered_and_placeless(bookings, False)
    assert covered == {"2026-09-24"}
    assert placeless == {"2026-09-21"}


def test_placeless_counts_as_covered_when_flag_on():
    bookings = [{"day": "2026-09-21T10:00:00.000Z", "state": "ACTIVE", "place": None}]
    covered, placeless = planner.covered_and_placeless(bookings, True)
    assert covered == {"2026-09-21"} and placeless == {"2026-09-21"}


def test_candidates_keep_preference_order_and_ignore_non_pending_days():
    a = Combo(1, "A", 3, "Large")
    b = Combo(2, "B", 2, "Medium")
    found = planner.candidates_by_day(
        ["2026-09-24"],
        [(a, ["2026-09-24T10:00:00.000Z", "2026-09-25T10:00:00.000Z"]),
         (b, ["2026-09-24T10:00:00.000Z"])],
    )
    assert list(found) == ["2026-09-24"]
    assert [c.combo for c in found["2026-09-24"]] == [a, b]
    assert found["2026-09-24"][0].raw_day == "2026-09-24T10:00:00.000Z"


def test_jwt_expiry_and_time_left():
    now = datetime(2026, 9, 20, tzinfo=timezone.utc)
    token = make_token(now + timedelta(days=2))
    assert auth.jwt_expiry(token) == now + timedelta(days=2)
    assert auth.time_left(token, now) == timedelta(days=2)


def test_jwt_expiry_none_for_garbage():
    assert auth.jwt_expiry("not-a-jwt") is None


def test_normalize_token_strips_quotes_and_bearer():
    assert auth.normalize_token('"Bearer abc.def.ghi"') == "abc.def.ghi"
    assert auth.normalize_token("  ") is None
    assert auth.normalize_token(None) is None


def test_holiday_dates_merges_library_calendar_and_extra_dates():
    skip = planner.holiday_dates("ES", "CT", ["2026-09-24"], date(2026, 9, 20), 30)
    assert "2026-09-24" in skip  # La Mercè: owner-supplied, the library lacks it
    assert "2026-10-12" in skip  # national holiday from the library
    assert "2026-09-21" not in skip


def test_holiday_dates_covers_year_rollover():
    skip = planner.holiday_dates("ES", "CT", [], date(2026, 12, 20), 30)
    assert {"2026-12-25", "2027-01-01"} <= skip


def test_bookable_until_moves_a_week_at_the_thursday_release():
    tz = ZoneInfo("Europe/Madrid")
    release = time(16, 0)
    thursday = datetime(2026, 9, 24, 15, 59, 59, tzinfo=tz)
    assert planner.bookable_until(thursday, release) == date(2026, 9, 27)
    assert planner.bookable_until(thursday.replace(hour=16, second=0), release) == date(2026, 10, 4)
    assert planner.bookable_until(datetime(2026, 9, 21, 9, tzinfo=tz), release) == date(2026, 9, 27)
    assert planner.bookable_until(datetime(2026, 9, 27, 23, tzinfo=tz), release) == date(2026, 10, 4)


def test_next_release_is_this_thursday_until_it_passes():
    tz = ZoneInfo("Europe/Madrid")
    release = time(16, 0)
    assert planner.next_release(datetime(2026, 9, 21, 9, tzinfo=tz), release) == \
        datetime(2026, 9, 24, 16, tzinfo=tz)
    assert planner.next_release(datetime(2026, 9, 24, 16, 5, tzinfo=tz), release) == \
        datetime(2026, 10, 1, 16, tzinfo=tz)
    assert planner.next_release(datetime(2026, 9, 26, 12, tzinfo=tz), release) == \
        datetime(2026, 10, 1, 16, tzinfo=tz)
