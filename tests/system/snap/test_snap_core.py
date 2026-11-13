"""Offline tests for the region-screenshot tool's naming, crop maths, hotkey parsing and config."""

import json
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "system" / "snap"))

from snap_core import (  # noqa: E402
    MIN_SELECTION_SIDE,
    MOD_CONTROL,
    MOD_SHIFT,
    ConfigError,
    Rect,
    load_config,
    parse_config,
    parse_hotkey,
    screen_to_image,
    selection_rect,
    unique_path,
    window_geometry,
)


class NamingTests(unittest.TestCase):
    def test_name_carries_date_time_and_milliseconds(self):
        moment = datetime(2026, 3, 7, 9, 5, 2, 42_000)
        with tempfile.TemporaryDirectory() as folder:
            self.assertEqual(unique_path(Path(folder), moment).name, "snap_2026-03-07_09-05-02-042.jpg")

    def test_two_captures_in_the_same_millisecond_never_collide(self):
        moment = datetime(2026, 3, 7, 9, 5, 2, 42_000)
        with tempfile.TemporaryDirectory() as folder:
            names = []
            for _ in range(3):
                path = unique_path(Path(folder), moment)
                path.write_bytes(b"x")
                names.append(path.name)
        self.assertEqual(len(set(names)), 3)
        self.assertEqual(names[1], "snap_2026-03-07_09-05-02-042_2.jpg")

    def test_captures_in_the_same_second_get_different_names(self):
        with tempfile.TemporaryDirectory() as folder:
            first = unique_path(Path(folder), datetime(2026, 3, 7, 9, 5, 2, 100_000))
            second = unique_path(Path(folder), datetime(2026, 3, 7, 9, 5, 2, 600_000))
        self.assertNotEqual(first, second)


class CropMathsTests(unittest.TestCase):
    def test_drag_in_any_direction_gives_the_same_rect(self):
        size = (1920, 1080)
        forward = selection_rect((100, 200), (400, 500), size)
        backward = selection_rect((400, 500), (100, 200), size)
        self.assertEqual(forward, Rect(100, 200, 400, 500))
        self.assertEqual(backward, forward)

    def test_selection_is_clamped_to_the_image(self):
        self.assertEqual(selection_rect((-50, -50), (3000, 2000), (1920, 1080)), Rect(0, 0, 1920, 1080))

    def test_a_click_or_tiny_drag_is_not_a_capture(self):
        self.assertIsNone(selection_rect((10, 10), (10, 10), (100, 100)))
        edge = MIN_SELECTION_SIDE - 1
        self.assertIsNone(selection_rect((10, 10), (10 + edge, 80), (100, 100)))
        self.assertIsNotNone(selection_rect((10, 10), (10 + MIN_SELECTION_SIDE, 80), (100, 100)))

    def test_secondary_monitor_left_of_primary_has_a_negative_origin(self):
        # Virtual screen starts at x=-1920: a point on the left monitor maps to a small image x.
        self.assertEqual(screen_to_image(-1820, 40, (-1920, 0)), (100, 40))
        # The same maths puts the primary monitor's origin at image x=1920.
        self.assertEqual(screen_to_image(0, 0, (-1920, 0)), (1920, 0))

    def test_monitor_above_the_primary_offsets_y(self):
        self.assertEqual(screen_to_image(500, -300, (0, -1080)), (500, 780))

    def test_drag_across_monitors_crops_in_image_space(self):
        origin = (-1920, 0)
        start = screen_to_image(-100, 100, origin)
        end = screen_to_image(300, 400, origin)
        self.assertEqual(selection_rect(start, end, (3840, 1080)), Rect(1820, 100, 2220, 400))

    def test_scaled_display_uses_physical_pixels_unchanged(self):
        # At 150% the process is DPI aware, so screen and image units are both physical pixels.
        self.assertEqual(screen_to_image(2400, 1350, (0, 0)), (2400, 1350))
        self.assertEqual(selection_rect((0, 0), (2400, 1350), (2880, 1620)), Rect(0, 0, 2400, 1350))

    def test_window_geometry_signs_negative_and_positive_origins(self):
        self.assertEqual(window_geometry(-1920, 0, 3840, 1080), "3840x1080-1920+0")
        self.assertEqual(window_geometry(0, -1080, 1920, 2160), "1920x2160+0-1080")


class HotkeyTests(unittest.TestCase):
    def test_default_hotkey(self):
        self.assertEqual(parse_hotkey("ctrl+shift+a"), (MOD_CONTROL | MOD_SHIFT, ord("A")))

    def test_function_key_and_case_insensitivity(self):
        self.assertEqual(parse_hotkey("Ctrl+F12"), (MOD_CONTROL, 0x7B))

    def test_a_bare_key_is_rejected(self):
        with self.assertRaises(ConfigError):
            parse_hotkey("a")

    def test_unknown_modifier_and_key_are_rejected_by_name(self):
        with self.assertRaisesRegex(ConfigError, "'hyper'"):
            parse_hotkey("hyper+a")
        with self.assertRaisesRegex(ConfigError, "'enter'"):
            parse_hotkey("ctrl+enter")


class ConfigTests(unittest.TestCase):
    def test_empty_config_uses_defaults(self):
        config = parse_config({})
        self.assertEqual(config.hotkey, "ctrl+shift+a")
        self.assertEqual(config.jpeg_quality, 90)
        self.assertTrue(config.copy_to_clipboard)

    def test_values_are_read(self):
        config = parse_config(
            {"folder": "~/shots", "hotkey": "Ctrl+Alt+S", "jpeg_quality": 75, "copy_to_clipboard": False}
        )
        self.assertEqual(config.folder, Path("~/shots").expanduser())
        self.assertEqual(config.hotkey, "ctrl+alt+s")
        self.assertEqual(config.jpeg_quality, 75)
        self.assertFalse(config.copy_to_clipboard)

    def test_invalid_values_name_the_setting(self):
        for raw, needle in (
            ({"jpeg_quality": 0}, "jpeg_quality"),
            ({"jpeg_quality": 101}, "jpeg_quality"),
            ({"jpeg_quality": True}, "jpeg_quality"),
            ({"jpeg_quality": "90"}, "jpeg_quality"),
            ({"copy_to_clipboard": "yes"}, "copy_to_clipboard"),
            ({"folder": ""}, "folder"),
            ({"hotkey": 5}, "hotkey"),
            ({"jpg_quality": 80}, "jpg_quality"),
        ):
            with self.subTest(raw=raw):
                with self.assertRaisesRegex(ConfigError, needle):
                    parse_config(raw)

    def test_missing_file_means_defaults(self):
        with tempfile.TemporaryDirectory() as folder:
            self.assertEqual(load_config(Path(folder) / "absent.json"), parse_config({}))

    def test_file_round_trip_and_bad_json(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "config.json"
            path.write_text(json.dumps({"jpeg_quality": 60}), encoding="utf-8")
            self.assertEqual(load_config(path).jpeg_quality, 60)
            path.write_text("{not json", encoding="utf-8")
            with self.assertRaises(ConfigError):
                load_config(path)
            path.write_text("[1, 2]", encoding="utf-8")
            with self.assertRaises(ConfigError):
                load_config(path)

    def test_sample_config_is_valid(self):
        sample = Path(__file__).resolve().parents[3] / "system" / "snap" / "config.sample.json"
        self.assertEqual(load_config(sample).hotkey, "ctrl+shift+a")


if __name__ == "__main__":
    unittest.main()
