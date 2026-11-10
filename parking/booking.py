"""Which days to book, and booking one day: each free slot in preference order until one sticks.

Shared by the regular poll and the Thursday burst workers. The site answers a
`createBooking` for a slot that has just gone with a booking in state REJECTED
(not an error), so a REJECTED answer moves straight on to the next slot
instead of stopping the day's attempt there.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from datetime import datetime, time
from enum import StrEnum
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


class DayStatus(StrEnum):
    """How one day's booking attempt ended. The values are the JSONL wire words of a burst worker."""

    BOOKED = "booked"  # confirmed ACTIVE on the site
    WOULD_BOOK = "would-book"  # dry-run: a slot was free, nothing was sent
    FAILED = "failed"  # slots were free but every booking was refused or errored
    UNCONFIRMED = "unconfirmed"  # a booking was sent and the site never showed it ACTIVE
    NO_SLOT = "no-slot"  # nothing free for the day
    COVERED = "covered"  # the day already had an ACTIVE booking (burst worker only)
    ERROR = "error"  # the attempt itself broke: auth, rate limit, server, or a worker that died
    UNKNOWN = "unknown"  # no result arrived in time; never counted as booked or as nothing free

    @property
    def is_covered(self) -> bool:
        """The day ends with a parking: booked now or already booked."""
        return self in (DayStatus.BOOKED, DayStatus.COVERED)

    @property
    def is_broken(self) -> bool:
        """The attempt did not run to an answer, so the run it belongs to is an error."""
        return self in (DayStatus.ERROR, DayStatus.UNKNOWN)


class RunStatus(StrEnum):
    """How one poll or burst ended."""

    INACTIVE = "inactive"  # polling is off in this time window
    BACKOFF = "backoff"  # not due yet
    NO_TOKEN = "no-token"  # no login token in the environment
    TOKEN_EXPIRED = "token-expired"  # the login token is past its expiry
    NOTHING_PENDING = "nothing-pending"  # every released target day already has a booking
    NO_SLOTS = "no-slots"  # days are open but no day had a free slot
    CHECKED = "checked"  # at least one day was tried; the per-day lists say how each ended
    ERROR = "error"  # the run broke before or while trying


@dataclass
class RunResult:
    """How one poll or burst ended; `status` is closed, so a word outside `RunStatus` is refused."""

    status: RunStatus
    pending: List[str] = field(default_factory=list)
    booked: List[str] = field(default_factory=list)
    would_book: List[str] = field(default_factory=list)
    failed: List[str] = field(default_factory=list)  # days where slots were free but booking failed
    unconfirmed: List[str] = field(default_factory=list)  # days where a sent booking was never confirmed
    error: str = ""

    def __post_init__(self) -> None:
        self.status = RunStatus(self.status)


@dataclass(frozen=True)
class DayOutcome:
    day: str
    status: DayStatus  # closed: a word outside `DayStatus` is refused
    label: str = ""  # the slot booked / would book / left unconfirmed
    error: str = ""  # why the last attempt failed, for the booking-failed alert

    def __post_init__(self) -> None:
        object.__setattr__(self, "status", DayStatus(self.status))


def derive_run_status(outcomes: Sequence[DayOutcome]) -> RunStatus:
    """The one mapping from per-day outcomes to how the run ended.

    A broken day (`error`/`unknown`) makes the run an error; every day without a
    free slot is `no-slots`; anything else tried at least one day, so `checked`.
    No outcomes means no day was tried: `nothing-pending`, never `no-slots`.
    """
    if not outcomes:
        return RunStatus.NOTHING_PENDING
    if any(o.status.is_broken for o in outcomes):
        return RunStatus.ERROR
    if all(o.status == DayStatus.NO_SLOT for o in outcomes):
        return RunStatus.NO_SLOTS
    return RunStatus.CHECKED


def result_from_outcomes(pending: List[str], outcomes: Sequence[DayOutcome]) -> RunResult:
    """The run result for `pending` days once every one was tried: status from `derive_run_status`."""
    def days(status: DayStatus) -> List[str]:
        return [o.day for o in outcomes if o.status == status]

    broken = "; ".join(f"{o.day}: {o.error}" for o in outcomes if o.status.is_broken)
    return RunResult(derive_run_status(outcomes), pending=pending, booked=days(DayStatus.BOOKED),
                     would_book=days(DayStatus.WOULD_BOOK), failed=days(DayStatus.FAILED),
                     unconfirmed=days(DayStatus.UNCONFIRMED), error=broken)


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
            return DayOutcome(day, DayStatus.WOULD_BOOK, label)
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
            return DayOutcome(day, DayStatus.BOOKED, label)
        if state == REJECTED:
            logger.warning("⚠️ %s was REJECTED by the site; trying the next slot", label)
            error = f"{label} was rejected"
            continue
        logger.error("❌ createBooking returned but %s is not confirmed in my bookings", day)
        return DayOutcome(day, DayStatus.UNCONFIRMED, label)
    return DayOutcome(day, DayStatus.FAILED, error=error) if error else DayOutcome(day, DayStatus.NO_SLOT)
