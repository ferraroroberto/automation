"""Pure counting helpers of the weekly Photos album: tile labels, per-day count checks, settling.

Split from ``photos_weekly_album`` so the browser flow stays small; nothing here touches a page.
"""

from __future__ import annotations

import re
from datetime import date, datetime, timedelta
from typing import Callable, Optional

LABEL_RE = re.compile(r"^(Photo|Video)\b.* - ([A-Z][a-z]{2} \d{1,2}, \d{4}, \d{1,2}:\d{2}:\d{2} [AP]M)$")


def label_datetime(label: str) -> Optional[datetime]:
    """Date taken from a tile label like ``Photo - Landscape - Oct 3, 2026, 8:23:55 AM``."""
    match = LABEL_RE.match(" ".join(label.split()))  # also folds the narrow no-break space before AM/PM
    return datetime.strptime(match.group(2), "%b %d, %Y, %I:%M:%S %p") if match else None


def label_date(label: str) -> Optional[date]:
    taken = label_datetime(label)
    return taken.date() if taken else None


def days_of(start: date, end: date) -> list[date]:
    return [start + timedelta(days=n) for n in range((end - start).days + 1)]


def count_mismatches(per_day: dict[date, int], tiles_per_day: dict[date, int]) -> list[str]:
    """One ``<day>: selected N, tiles M`` entry per day whose two counts disagree."""
    return [f"{day}: selected {per_day.get(day, 0)}, tiles {tiles_per_day.get(day, 0)}"
            for day in sorted(set(per_day) | set(tiles_per_day))
            if per_day.get(day, 0) != tiles_per_day.get(day, 0)]


def settled_value(read: Callable[[], int], pause: Callable[[], None], floor: int,
                  stable_reads: int = 3, max_reads: int = 40) -> int:
    """``read()`` once it has stopped changing for ``stable_reads`` reads in a row.

    Photos counts a big day's items up after a select-all click rather than jumping to the total, so
    a single read right after the click is short. A value still at ``floor`` (nothing selected yet)
    never counts as settled until the reads run out.
    """
    last, run_length = -1, 0
    for _ in range(max_reads):
        value = read()
        run_length = run_length + 1 if value == last else 1
        last = value
        if run_length >= stable_reads and value != floor:
            return value
        pause()
    return last
