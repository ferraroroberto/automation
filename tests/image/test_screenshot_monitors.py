"""Monitor selection in screenshot.py: bounds and the bottom-left display."""

import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "image"))

import screenshot  # noqa: E402

VIRTUAL = {"left": 0, "top": 0, "width": 3840, "height": 2160}
TOP_LEFT = {"left": 0, "top": 0, "width": 1920, "height": 1080}
BOTTOM_LEFT = {"left": 0, "top": 1080, "width": 1920, "height": 1080}
TOP_RIGHT = {"left": 1920, "top": 0, "width": 1920, "height": 1080}


class BottomLeftMonitorTests(unittest.TestCase):
    def test_picks_the_lowest_row_then_the_leftmost(self):
        monitors = [VIRTUAL, TOP_LEFT, TOP_RIGHT, BOTTOM_LEFT]
        self.assertIs(screenshot.get_bottom_left_monitor(monitors), BOTTOM_LEFT)

    def test_single_row_picks_the_leftmost(self):
        monitors = [VIRTUAL, TOP_RIGHT, TOP_LEFT]
        self.assertIs(screenshot.get_bottom_left_monitor(monitors), TOP_LEFT)


class MonitorIndexBoundsTests(unittest.TestCase):
    def test_index_one_past_the_last_monitor_is_a_value_error(self):
        fake = mock.MagicMock()
        fake.__enter__.return_value.monitors = [VIRTUAL, TOP_LEFT, TOP_RIGHT]
        with tempfile.TemporaryDirectory() as tmp, mock.patch.object(screenshot, "mss", return_value=fake):
            with self.assertRaises(ValueError):
                screenshot.capture_specific_monitor(3, tmp)


if __name__ == "__main__":
    unittest.main()
