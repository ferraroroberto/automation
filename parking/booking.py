"""Which days to book, and booking one day: each free slot in preference order until one sticks.

Shared by the regular poll and the Thursday burst workers. The site answers a
`createBooking` for a slot that has just gone with a booking in state REJECTED
(not an error), so a REJECTED answer moves straight on to the next slot - the
2026-09-24 burst lost two days by stopping there (#139).
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from datetime import datetime, time
from typing import Callable, Iterable, List, Mapping, Optional, Sequence, Set

from parking import planner
from parking.api import ApiError, AuthError, ParkingApi, RateLimited, ServerError
from parking.config import Config

logger = logging.getLogger("parking.booking")

ACTIVE = "ACTIVE"
REJECTED = "REJECTED"
DEFAULT_RELEASE = time(16, 0)  # the site's weekly release: Thursday 16:00


def release_time(cfg: Config) -> time:
    return cfg.thursday_burst.target if cfg.thursday_burst else DEFAULT_RELEASE


def combos(cfg: Config) -> List[planner.Combo]:
    """Every centre x size, in preference (config) order."""
    return [planner.Combo(c.id, c.name, s.id, s.name) for c in cfg.centers for s in cfg.sizes]


def pending_days(cfg: Config, bookings: Sequence[Mapping], now: datetime,
                 released_at: Optional[datetime] = None) -> List[str]:
    """Target days already released by `released_at` (default `now`) that have no ACTIVE booking.

    Days not yet released (beyond next Sunday, or beyond this Sunday before the
    Thursday release) are left out: nothing can be booked there, so they must
    neither keep the regular poll busy nor stop the burst from finishing early.
    """
    covered, placeless = planner.covered_and_placeless(bookings, cfg.treat_placeless_as_covered)
    if placeless:
        logger.info("ℹ️ days with an ACTIVE booking but no place assigned: %s (covered=%s)",
                    sorted(placeless), cfg.treat_placeless_as_covered)
    last = planner.bookable_until(released_at or now, release_time(cfg)).isoformat()
    released = [d for d in planner.target_dates(now.date(), cfg.weekdays, cfg.horizon_days) if d <= last]
    holidays = planner.holiday_dates(cfg.holiday_country, cfg.holiday_subdiv, cfg.skip_dates,
                                     now.date(), cfg.horizon_days)
    targets = [d for d in released if d not in holidays]
    if len(targets) < len(released):
        logger.info("ℹ️ skipping holiday(s): %s", [d for d in released if d in holidays])
    return planner.pending_dates(targets, covered)


@dataclass
class RunResult:
    """How one poll or burst ended."""

    status: str  # inactive | backoff | no-token | token-expired | nothing-pending | no-slots | checked | error
    pending: List[str] = field(default_factory=list)
    booked: List[str] = field(default_factory=list)
    would_book: List[str] = field(default_factory=list)
    error: str = ""


@dataclass(frozen=True)
class DayOutcome:
    day: str
    status: str  # booked | would-book | failed | unconfirmed | no-slot | covered | error | unknown
    label: str = ""  # the slot booked / would book / left unconfirmed
    error: str = ""  # why the last attempt failed, for the booking-failed alert


def slot_label(candidate: planner.Candidate) -> str:
    return f"{candidate.day} at {candidate.combo.center_name} ({candidate.combo.size_name})"


def _describe(records: Sequence[Mapping]) -> str:
    if not records:
        return "no booking returned"
    return "; ".join(f"id={r.get('id')} state={r.get('state')} place={(r.get('place') or {}).get('place')}"
                     for r in records)


def _state_for_day(records: Iterable[Mapping], day: str, ids: Optional[Set[str]] = None) -> Optional[str]:
    """ACTIVE if any booking for `day` is active; else the state of the last one
    among `ids` (this attempt's own bookings, when given); else None."""
    mine = [r for r in records if planner.day_key(str(r.get("day") or "")) == day]
    if any(r.get("state") == ACTIVE for r in mine):
        return ACTIVE
    if ids is not None:
        mine = [r for r in mine if str(r.get("id")) in ids]
    return mine[-1].get("state") if mine else None


def book_day(api: ParkingApi, day: str, candidates: List[planner.Candidate], plate: str, type_: str,
             dry_run: bool, say: Callable[[str], object] = lambda _text: None) -> DayOutcome:
    """Try `candidates` (one day, preference order) until one is confirmed ACTIVE.

    One ACTIVE booking at most: an answer that is neither ACTIVE nor REJECTED is
    re-checked against my bookings, and if that is still inconclusive the day
    stops there ("unconfirmed") rather than risk a second booking.
    Auth/rate-limit/server errors propagate; any other API error tries the next slot.
    """
    error = ""
    for candidate in candidates:
        label = slot_label(candidate)
        if dry_run:
            logger.info("ℹ️ [dry-run] would book %s", label)
            return DayOutcome(day, "would-book", label)
        say(f"Booking {label}...")
        try:
            created = api.create_booking(candidate.combo.center_id, candidate.combo.size_id, type_,
                                         plate, candidate.raw_day)
        except ApiError as exc:
            if isinstance(exc, (AuthError, RateLimited, ServerError)):
                raise
            logger.warning("⚠️ booking %s failed (%s); trying the next slot", label, exc)
            error = str(exc)
            continue
        logger.info("ℹ️ createBooking %s -> %s", label, _describe(created))
        state = _state_for_day(created, day)
        if state not in (ACTIVE, REJECTED):
            # An earlier REJECTED record for the same day must not decide this attempt.
            ids = {str(r["id"]) for r in created if r.get("id") is not None}
            state = _state_for_day(api.my_bookings(), day, ids)
        if state == ACTIVE:
            logger.info("✅ booked %s", label)
            return DayOutcome(day, "booked", label)
        if state == REJECTED:
            logger.warning("⚠️ %s was REJECTED by the site; trying the next slot", label)
            error = f"{label} was rejected"
            continue
        logger.error("❌ createBooking returned but %s is not confirmed in my bookings", day)
        return DayOutcome(day, "unconfirmed", label)
    return DayOutcome(day, "failed", error=error) if error else DayOutcome(day, "no-slot")
