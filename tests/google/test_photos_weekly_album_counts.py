"""Offline tests for the Photos weekly album's per-day counting and its completeness guard
(no browser, no network)."""

import unittest
from datetime import date

from google import photos_weekly_album as pwa


class SettledValueTests(unittest.TestCase):
    """A big day's select-all counts up over a few reads; a single early read is short."""

    @staticmethod
    def _reads(values):
        feed = iter(values)
        return lambda: next(feed)

    def test_a_counter_still_climbing_is_waited_out(self):
        reads = self._reads([100, 180, 260, 290, 290, 290, 290])
        self.assertEqual(pwa.settled_value(reads, lambda: None, floor=100), 290)

    def test_a_value_stuck_at_the_floor_is_not_settled(self):
        pauses = []
        value = pwa.settled_value(lambda: 100, lambda: pauses.append(1), floor=100, max_reads=5)
        self.assertEqual((value, len(pauses)), (100, 5))

    def test_a_small_day_settles_after_the_stable_reads(self):
        pauses = []
        value = pwa.settled_value(self._reads([115, 115, 115, 115]), lambda: pauses.append(1), floor=100)
        self.assertEqual((value, len(pauses)), (115, 2))


class CountMismatchTests(unittest.TestCase):
    SAT, SUN = date(2026, 10, 24), date(2026, 10, 25)

    def test_equal_counts_report_nothing(self):
        self.assertEqual(pwa.count_mismatches({self.SAT: 290, self.SUN: 40}, {self.SAT: 290, self.SUN: 40}), [])

    def test_the_short_day_is_named(self):
        got = pwa.count_mismatches({self.SAT: 279, self.SUN: 40}, {self.SAT: 290, self.SUN: 40})
        self.assertEqual(got, ["2026-10-24: selected 279, tiles 290"])

    def test_a_day_only_the_sweep_saw_is_named(self):
        got = pwa.count_mismatches({self.SAT: 5}, {self.SAT: 5, self.SUN: 3})
        self.assertEqual(got, ["2026-10-25: selected 0, tiles 3"])


class CollectGuardTests(unittest.TestCase):
    START, END = date(2026, 10, 24), date(2026, 10, 30)

    def _web(self, per_day, tiles_per_day, excluded=0, selected=0):
        web = pwa.PhotosWeb.__new__(pwa.PhotosWeb)
        web.screenshot_labels = lambda s, e: {}
        web.select_week = lambda s, e: per_day
        web.drop_excluded = lambda s, e, x: (tiles_per_day, {}, excluded, 0)
        web._counter = lambda: selected
        return web

    def test_a_short_day_aborts_and_names_the_day(self):
        sat = self.START
        web = self._web({sat: 279}, {sat: 290})
        with self.assertRaises(pwa.FlowError) as ctx:
            web.collect(self.START, self.END)
        self.assertIn("2026-10-24: selected 279, tiles 290", str(ctx.exception))

    def test_tiles_the_headers_never_selected_abort(self):
        sat, sun = self.START, date(2026, 10, 25)
        # The sweep saw 11 tiles the header's select-all left unticked; the counter confirms it.
        web = self._web({sat: 289, sun: 40}, {sat: 300, sun: 40}, excluded=4, selected=325)
        with self.assertRaises(pwa.FlowError) as ctx:
            web.collect(self.START, self.END)
        self.assertIn("2026-10-24: selected 289, tiles 300", str(ctx.exception))

    def test_matching_counts_pass(self):
        sat = self.START
        sel = self._web({sat: 290}, {sat: 290}, excluded=4, selected=286).collect(self.START, self.END)
        self.assertEqual((sel.week_tiles, sel.day_total, sel.selected), (290, 290, 286))


class _LazyTimeline(pwa.PhotosWeb):
    """A timeline that renders ``VIEW`` tiles; scrolling renders and loads more.

    Like Photos, a day header's select-all ticks only that day's tiles loaded so far (up to
    ``AHEAD`` below the view), so a day taller than that is selected short.
    """

    VIEW, AHEAD, STEP = 20, 30, 10

    def __init__(self, days):
        self.tiles = [(f"t{d:%m%d}-{n}", d) for d, count in days for n in range(count)]
        self.selected: set[str] = set()
        self.pos = self.loaded = 0
        self.page = self

    def _move(self, pos):
        self.pos = pos
        self.loaded = max(self.loaded, pos + self.VIEW + self.AHEAD)

    @staticmethod
    def _label(day):
        return f"Photo - Landscape - {day:%b} {day.day}, {day.year}, 10:00:00 AM"

    def _scan(self):
        shown = self.tiles[self.pos:self.pos + self.VIEW]
        first = {}
        for tid, day in self.tiles:
            first.setdefault(day, tid)
        return {"visible": True,
                "headers": [{"key": tid, "label": self._label(day), "checked": False}
                            for tid, day in shown if first[day] == tid],
                "tiles": [{"id": tid, "label": self._label(day), "checked": tid in self.selected, "box": True}
                          for tid, day in shown]}

    def _open(self, url):
        self._move(0)

    def _wheel(self):
        if self.pos + self.VIEW >= len(self.tiles):
            return False
        self._move(self.pos + self.STEP)
        return True

    def _counter(self):
        return len(self.selected)

    def _settled_counter(self, before):
        return self._counter()

    def screenshot_labels(self, start, end):
        return {}

    def _click_checked(self, selector, reveal, want, what):
        key = selector.split('"')[1]
        if selector.startswith("[data-pwa-tile"):
            (self.selected.add if want else self.selected.discard)(key)
            return
        day = dict(self.tiles)[key]
        self.selected |= {tid for i, (tid, d) in enumerate(self.tiles) if d == day and i < self.loaded}

    def wait_for_timeout(self, ms):
        pass

    def evaluate(self, js, arg=None):
        if js == pwa.SCROLL_POS_JS:
            return self.pos
        if js == pwa.SCROLL_TOP_JS:
            return self._move(arg)
        if js == pwa.HEAD_NEAR_TOP_JS:
            index = [tid for tid, _ in self.tiles].index(arg)
            self._move(index)
            return True
        return self._scan()


class LazyDayTests(unittest.TestCase):
    def test_tiles_a_short_day_header_left_are_ticked_and_the_week_is_whole(self):
        start, end = date(2026, 10, 24), date(2026, 10, 30)
        counts = {date(2026, 10, 30): 42, date(2026, 10, 29): 74, date(2026, 10, 28): 31,
                  date(2026, 10, 27): 29, date(2026, 10, 26): 25, date(2026, 10, 25): 28, start: 300}
        web = _LazyTimeline([*counts.items(), (date(2026, 10, 23), 30)])
        sel = web.collect(start, end)
        week = {tid for tid, day in web.tiles if start <= day <= end}
        self.assertEqual((sel.selected_per_day, sel.selected, web.selected), (counts, 529, week))
        self.assertLess(sel.per_day[start], 300)  # the header alone fell short, as on the live library


if __name__ == "__main__":
    unittest.main()
