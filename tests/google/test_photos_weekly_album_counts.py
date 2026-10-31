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
        self.assertEqual(got, ["2026-10-24: headers 279, tiles 290"])

    def test_a_day_only_the_sweep_saw_is_named(self):
        got = pwa.count_mismatches({self.SAT: 5}, {self.SAT: 5, self.SUN: 3})
        self.assertEqual(got, ["2026-10-25: headers 0, tiles 3"])


class CollectGuardTests(unittest.TestCase):
    START, END = date(2026, 10, 24), date(2026, 10, 30)

    def _web(self, per_day, tiles_per_day, excluded=0, selected=0):
        web = pwa.PhotosWeb.__new__(pwa.PhotosWeb)
        web.screenshot_labels = lambda s, e: {}
        web.select_week = lambda s, e: per_day
        web.drop_excluded = lambda s, e, x: (tiles_per_day, excluded, 0)
        web._counter = lambda: selected
        return web

    def test_a_short_day_aborts_and_names_the_day(self):
        sat = self.START
        web = self._web({sat: 279}, {sat: 290})
        with self.assertRaises(pwa.FlowError) as ctx:
            web.collect(self.START, self.END)
        self.assertIn("2026-10-24: headers 279, tiles 290", str(ctx.exception))

    def test_an_early_read_of_the_oldest_day_is_vouched_for_by_the_final_counter(self):
        sat, sun = self.START, date(2026, 10, 25)
        # Oldest day read 289 of its 300 tiles; 4 screenshots unticked; counter holds 300 + 40 - 4.
        sel = self._web({sat: 289, sun: 40}, {sat: 300, sun: 40}, excluded=4, selected=336).collect(
            self.START, self.END)
        self.assertEqual((sel.per_day[sat], sel.day_total, sel.week_tiles), (300, 340, 340))

    def test_extra_tiles_on_the_oldest_day_abort_when_the_counter_disagrees(self):
        sat, sun = self.START, date(2026, 10, 25)
        # The sweep counted 11 tiles the headers never selected; the counter confirms the headers.
        web = self._web({sat: 289, sun: 40}, {sat: 300, sun: 40}, excluded=4, selected=325)
        with self.assertRaises(pwa.FlowError) as ctx:
            web.collect(self.START, self.END)
        self.assertIn("2026-10-24: headers 289, tiles 300", str(ctx.exception))

    def test_a_short_day_after_the_oldest_is_never_reconciled(self):
        sat, sun = self.START, date(2026, 10, 25)
        self.assertEqual(pwa.reconcile_boundary_day({sat: 289, sun: 30}, {sat: 300, sun: 40}, sat, 336, 4),
                         {sat: 289, sun: 30})

    def test_matching_counts_pass(self):
        sat = self.START
        sel = self._web({sat: 290}, {sat: 290}, excluded=4, selected=286).collect(self.START, self.END)
        self.assertEqual((sel.week_tiles, sel.day_total, sel.selected), (290, 290, 286))


if __name__ == "__main__":
    unittest.main()
