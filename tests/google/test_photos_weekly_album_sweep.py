"""Offline tests for the weekly Photos album's tile sweep and for clearing a failed collect's
selection (tools/google/photos_weekly_album.py; no browser, no network)."""

import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest import mock

from tools.google import photos_weekly_album as pwa

START, END = date(2026, 11, 7), date(2026, 11, 13)


class _GridPage:
    def evaluate(self, script, *args):
        return None

    def wait_for_timeout(self, ms):
        pass


class TileSweepTests(unittest.TestCase):
    def test_the_sweep_counts_a_burst_stack_its_day_header_selects(self):
        """A burst stack is one item to the day header, so the sweep must date its tile too."""
        tiles = [{"id": i, "label": label, "checked": True, "box": True} for i, label in [
            ("a", "Photo - Landscape - Nov 7, 2026, 10:09:12 PM"),
            ("b", "Burst photo - Landscape - Nov 7, 2026, 10:09:43 PM - 3 photos in sequence"),
            ("c", "Photo - Portrait - Nov 6, 2026, 8:00:00 PM"),
        ]]
        web = pwa.PhotosWeb(Path("unused"))
        web.page = _GridPage()
        with mock.patch.object(web, "_scan", return_value={"tiles": tiles}), \
                mock.patch.object(web, "_wheel", return_value=True):
            sweep = web.drop_excluded(START, END, {})
        self.assertEqual(sweep.tiles_per_day, {date(2026, 11, 7): 2})


def _tile(tile_id, label, head=None):
    return {"id": tile_id, "label": label, "checked": True, "box": True, "head": head}


class DayHeaderTests(unittest.TestCase):
    """A tile counts under the day header it sits under, not under its label's date."""

    def _sweep(self, scans):
        web = pwa.PhotosWeb(Path("unused"))
        web.page = _GridPage()
        with mock.patch.object(web, "_scan", side_effect=scans * pwa.SWEEP_MAX_PASSES), \
                mock.patch.object(web, "_advance", side_effect=([True] * (len(scans) - 1) + [False]) * pwa.SWEEP_MAX_PASSES), \
                self.assertLogs("photos_weekly_album", "INFO") as logs:
            sweep = web.drop_excluded(START, END, {})
        return sweep, "\n".join(logs.output)

    def test_a_tile_dated_the_day_before_the_week_counts_under_the_first_day_header(self):
        """The first day's header selects an item taken near midnight whose label reads the day before."""
        scans = [{"headers": [{"key": "a", "label": "Photo - Landscape - Nov 7, 2026, 10:09:12 PM"}],
                  "tiles": [_tile("a", "Photo - Landscape - Nov 7, 2026, 10:09:12 PM", "a"),
                            _tile("b", "Photo - Landscape - Nov 6, 2026, 11:58:00 PM", "a")]},
                 # the next scan has scrolled past the header: "b" carries the header it was seen under
                 {"headers": [{"key": "c", "label": "Photo - Portrait - Nov 6, 2026, 8:00:00 PM"}],
                  "tiles": [_tile("b", "Photo - Landscape - Nov 6, 2026, 11:58:00 PM"),
                            _tile("c", "Photo - Portrait - Nov 6, 2026, 8:00:00 PM", "c")]}]
        sweep, log = self._sweep(scans)
        self.assertEqual(sweep.tiles_per_day, {date(2026, 11, 7): 2})
        self.assertIn("1 tile(s) whose label date differs from the day header they sit under "
                      "(label earlier 1; across the week edge 1)", log)

    def test_a_tile_carries_the_header_of_the_tile_before_it_in_the_scan(self):
        scans = [{"headers": [{"key": "a", "label": "Photo - Landscape - Nov 13, 2026, 12:30:00 AM"}],
                  "tiles": [_tile("a", "Photo - Landscape - Nov 13, 2026, 12:30:00 AM", "a")]},
                 {"headers": [],
                  "tiles": [_tile("a", "Photo - Landscape - Nov 13, 2026, 12:30:00 AM"),
                            _tile("z", "Photo - Landscape - Nov 14, 2026, 12:05:00 AM")]}]
        sweep, _ = self._sweep(scans)
        self.assertEqual(sweep.tiles_per_day, {date(2026, 11, 13): 2})

    def test_a_week_tile_under_a_header_outside_the_week_still_counts_by_its_label(self):
        """No week header selects it, so it is ticked by hand and counted on its label's day, as before."""
        scans = [{"headers": [{"key": "a", "label": "Photo - Landscape - Nov 14, 2026, 12:30:00 AM"}],
                  "tiles": [_tile("a", "Photo - Landscape - Nov 14, 2026, 12:30:00 AM", "a"),
                            _tile("b", "Photo - Landscape - Nov 13, 2026, 11:50:00 PM", "a")]}]
        sweep, _ = self._sweep(scans)
        self.assertEqual(sweep.tiles_per_day, {date(2026, 11, 13): 1})

    def test_tiles_whose_label_matches_their_header_log_no_disagreement(self):
        scans = [{"headers": [{"key": "a", "label": "Photo - Landscape - Nov 8, 2026, 9:00:00 AM"}],
                  "tiles": [_tile("a", "Photo - Landscape - Nov 8, 2026, 9:00:00 AM", "a"),
                            _tile("b", "Video - Landscape - Nov 8, 2026, 8:00:00 AM", "a")]}]
        sweep, log = self._sweep(scans)
        self.assertEqual(sweep.tiles_per_day, {date(2026, 11, 8): 2})
        self.assertNotIn("differs from the day header", log)


class SweepCapTests(unittest.TestCase):
    def test_a_sweep_that_keeps_finding_tiles_is_capped(self):
        scans = iter(range(1000))

        def scan():
            n = next(scans)
            return {"tiles": [{"id": f"x{n}", "label": "Photo - Landscape - Nov 7, 2026, 10:09:12 PM",
                               "checked": True, "box": True}]}

        web = pwa.PhotosWeb(Path("unused"))
        web.page = _GridPage()
        with mock.patch.object(web, "_scan", side_effect=scan), \
                mock.patch.object(web, "_advance", side_effect=[True, False] * pwa.SWEEP_MAX_PASSES), \
                self.assertLogs("photos_weekly_album", "WARNING"):
            web.drop_excluded(START, END, {})


class _FailingWeb:
    """A browser session whose collect fails after the day headers selected items."""

    instance = None

    def __init__(self, profile_dir):
        _FailingWeb.instance = self
        self.cleared = False

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return None

    def require_signed_in(self):
        pass

    def find_album(self, title):
        return None

    def collect(self, start, end):
        raise pwa.FlowError("saw 78 selectable tiles in the week but 197 were selected")

    def drop_selection(self):
        self.cleared = True


class FailedCollectTests(unittest.TestCase):
    def test_a_failed_collect_leaves_no_selection_behind(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg = pwa.Config(title_template="t", counters={}, profile_dir=Path(tmp),
                             log_file=Path(tmp) / "run.log")
            with mock.patch.object(pwa, "PhotosWeb", _FailingWeb), self.assertLogs("photos_weekly_album", "ERROR"):
                code = pwa.run(cfg, START, END, "t", dry_run=False)
        self.assertEqual(code, pwa.FAILURE_EXIT["error"])
        self.assertTrue(_FailingWeb.instance.cleared)


if __name__ == "__main__":
    unittest.main()
