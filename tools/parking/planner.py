"""Pure date-selection and slot-decision logic (no I/O, unit-tested)."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from typing import Dict, Iterable, List, Mapping, Sequence, Set, Tuple

import holidays

WEEKDAYS = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]


@dataclass(frozen=True)
class Combo:
    center_id: int
    center_name: str
    size_id: int
    size_name: str


@dataclass(frozen=True)
class Candidate:
    """One bookable (day, centre, size). `raw_day` is the value exactly as the API returned it."""

    day: str
    raw_day: str
    combo: Combo


def day_key(raw: str) -> str:
    """'2026-09-24T10:00:00.000Z' -> '2026-09-24'."""
    return raw[:10]


def holiday_dates(country: str, subdiv: str, extra: Iterable[str], today: date,
                  horizon_days: int) -> Set[str]:
    """Days we never book: the public-holiday calendar plus the owner's own `extra` dates."""
    years = range(today.year, (today + timedelta(days=horizon_days)).year + 1)
    calendar = holidays.country_holidays(country, subdiv=subdiv or None, years=years)
    return {d.isoformat() for d in calendar} | set(extra)


def release_moment(now: datetime, release: time) -> datetime:
    """This week's slot release (Thursday at `release`), in `now`'s timezone."""
    thursday = now.date() + timedelta(days=WEEKDAYS.index("thu") - now.weekday())
    return datetime.combine(thursday, release, tzinfo=now.tzinfo)


def next_release(now: datetime, release: time) -> datetime:
    """The first Thursday release at or after `now`."""
    this_week = release_moment(now, release)
    return this_week if now < this_week else this_week + timedelta(days=7)


def bookable_until(now: datetime, release: time) -> date:
    """Last day already released for booking: this week's Sunday until Thursday's
    release, then next week's Sunday (the release opens the following week)."""
    sunday = now.date() + timedelta(days=6 - now.weekday())
    return sunday if now < release_moment(now, release) else sunday + timedelta(days=7)


def target_dates(today: date, weekdays: Iterable[str], horizon_days: int) -> List[str]:
    wanted = {WEEKDAYS.index(w.strip().lower()[:3]) for w in weekdays}
    days = (today + timedelta(days=offset) for offset in range(horizon_days + 1))
    return [d.isoformat() for d in days if d.weekday() in wanted]


def covered_and_placeless(
    bookings: Sequence[Mapping], treat_placeless_as_covered: bool
) -> Tuple[Set[str], Set[str]]:
    """Days served by an ACTIVE booking, and days holding an ACTIVE booking with no place.

    A day is covered only if the booking has a place, unless `treat_placeless_as_covered`.
    """
    covered: Set[str] = set()
    placeless: Set[str] = set()
    for booking in bookings:
        if booking.get("state") != "ACTIVE" or not booking.get("day"):
            continue
        key = day_key(booking["day"])
        if (booking.get("place") or {}).get("place"):
            covered.add(key)
        else:
            placeless.add(key)
            if treat_placeless_as_covered:
                covered.add(key)
    return covered, placeless


def pending_dates(targets: Iterable[str], covered: Set[str]) -> List[str]:
    return [d for d in targets if d not in covered]


def candidates_by_day(
    pending: Iterable[str],
    slots: Iterable[Tuple[Combo, Sequence[str]]],
) -> Dict[str, List[Candidate]]:
    """Group free slots by pending day, keeping the caller's combo preference order."""
    wanted = set(pending)
    found: Dict[str, List[Candidate]] = {}
    for combo, raw_days in slots:
        for raw in raw_days:
            key = day_key(raw)
            if key in wanted:
                found.setdefault(key, []).append(Candidate(key, raw, combo))
    return found
