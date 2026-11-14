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
