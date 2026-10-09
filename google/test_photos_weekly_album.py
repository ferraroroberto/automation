"""Offline tests for google/photos_weekly_album.py and browser_stealth.py (no browser, no network)."""

import unittest
from datetime import date, datetime, timedelta
from pathlib import Path

import browser_stealth as bs
from google import photos_weekly_album as pwa

SAMPLE = Path(pwa.__file__).with_name("photos_weekly_album.json.sample")


class WeekRangeTests(unittest.TestCase):
    def test_saturday_run_takes_the_week_that_just_ended(self):
        self.assertEqual(pwa.week_range(date(2026, 10, 10)), (date(2026, 10, 3), date(2026, 10, 9)))

    def test_friday_is_still_open_so_the_previous_week_is_taken(self):
        self.assertEqual(pwa.week_range(date(2026, 10, 9)), (date(2026, 9, 26), date(2026, 10, 2)))

    def test_every_weekday_gives_a_saturday_to_friday_week_before_today(self):
        for offset in range(14):
            today = date(2026, 10, 10) + timedelta(days=offset)
            with self.subTest(today=today):
                start, end = pwa.week_range(today)
                self.assertEqual((start.weekday(), end.weekday()), (5, 4))
                self.assertEqual(end - start, timedelta(days=6))
                self.assertTrue(1 <= (today - end).days <= 7)

    def test_crosses_month_and_year_boundaries(self):
        self.assertEqual(pwa.week_range(date(2027, 1, 2)), (date(2026, 12, 26), date(2027, 1, 1)))

    def test_days_of_lists_both_ends(self):
        days = pwa.days_of(date(2026, 10, 3), date(2026, 10, 9))
        self.assertEqual((len(days), days[0], days[-1]), (7, date(2026, 10, 3), date(2026, 10, 9)))


class TitleTests(unittest.TestCase):
    COUNTERS = {"a": {"value": 10, "week_start": "2026-09-26"}, "b": {"value": 3, "week_start": "2026-09-26"}}

    def test_counters_go_up_one_per_week(self):
        self.assertEqual(pwa.counter_value(10, date(2026, 9, 26), date(2026, 9, 26)), 10)
        self.assertEqual(pwa.counter_value(10, date(2026, 9, 26), date(2026, 10, 3)), 11)
        self.assertEqual(pwa.counter_value(10, date(2026, 9, 26), date(2027, 9, 25)), 62)

    def test_weeks_before_the_base_count_down(self):
        self.assertEqual(pwa.counter_value(10, date(2026, 9, 26), date(2026, 9, 12)), 8)

    def test_base_week_must_start_on_saturday(self):
        with self.assertRaisesRegex(ValueError, "not a Saturday"):
            pwa.counter_value(1, date(2026, 9, 27), date(2026, 10, 3))

    def test_template_gets_every_counter(self):
        self.assertEqual(pwa.build_title("W{a}-{b}", self.COUNTERS, date(2026, 10, 10)), "W12-5")

    def test_unknown_placeholder_is_a_config_error(self):
        with self.assertRaisesRegex(ValueError, "placeholder 'c'"):
            pwa.build_title("W{a}-{c}", self.COUNTERS, date(2026, 10, 3))

    def test_committed_sample_builds_a_title(self):
        cfg = pwa.Config.load(SAMPLE)
        self.assertEqual(cfg.profile_dir, pwa.DEFAULT_PROFILE)
        self.assertIn("101", pwa.build_title(cfg.title_template, cfg.counters, date(2026, 1, 10)))


class LabelTests(unittest.TestCase):
    def test_photo_and_video_labels(self):
        self.assertEqual(pwa.label_datetime("Photo - Landscape - Oct 3, 2026, 8:23:55 AM"),
                         datetime(2026, 10, 3, 8, 23, 55))
        self.assertEqual(pwa.label_date("Video - Portrait - Oct 9, 2026, 6:47:52 PM"), date(2026, 10, 9))

    def test_narrow_no_break_space_before_am_pm(self):
        self.assertEqual(pwa.label_date("Photo - Portrait - Dec 31, 2026, 11:59:59\u202fPM"), date(2026, 12, 31))

    def test_unrecognised_labels_give_none(self):
        for label in ("", "Select", "Photo - Landscape - Oct 3, 8:23:55 AM", "Album - Oct 3, 2026, 8:23:55 AM"):
            with self.subTest(label=label):
                self.assertIsNone(pwa.label_date(label))


class ArgsTests(unittest.TestCase):
    def test_from_and_to_go_together_and_in_order(self):
        for argv in (["--from", "2026-10-03"], ["--from", "2026-10-09", "--to", "2026-10-03"],
                     ["--from", "2026-10-01", "--to", "2026-11-15"]):
            with self.subTest(argv=argv), self.assertRaises(SystemExit):
                pwa.parse_args(argv)

    def test_explicit_week(self):
        args = pwa.parse_args(["--from", "2026-10-03", "--to", "2026-10-09", "--dry-run"])
        self.assertEqual((args.start, args.end, args.dry_run), (date(2026, 10, 3), date(2026, 10, 9), True))


class _FakeCheckbox:
    """A checkbox that, like a day header's, has no size until the pointer is over its row."""

    def __init__(self, shown):
        self.shown, self.checked, self.events = shown, False, []

    def is_visible(self):
        return self.shown

    def hover(self):
        self.events.append("hover")
        self.shown = True

    def click(self):
        if not self.shown:
            raise TimeoutError("element is not visible")
        self.events.append("click")
        self.checked = not self.checked

    def wait_for(self, timeout):
        if not self.checked:
            raise TimeoutError("still unchecked")


class _FakePage:
    def __init__(self, box):
        self.box = box

    def locator(self, selector):
        return self

    @property
    def first(self):
        return self.box


class ClickCheckedTests(unittest.TestCase):
    def _click(self, box):
        web = pwa.PhotosWeb(Path("unused"))
        web.page = _FakePage(box)
        web._click_checked("[data-pwa-head=k]", "[data-pwa-head=k] >> xpath=..", True, "day")

    def test_hidden_checkbox_is_revealed_by_hovering_before_the_click(self):
        box = _FakeCheckbox(shown=False)
        self._click(box)
        self.assertEqual(box.events, ["hover", "click"])

    def test_shown_checkbox_is_clicked_without_hovering_what_it_covers(self):
        box = _FakeCheckbox(shown=True)
        self._click(box)
        self.assertEqual(box.events, ["click"])


class _FakeContext:
    def __init__(self):
        self.scripts = []

    def add_init_script(self, script):
        self.scripts.append(script)


class _FakePlaywright:
    def __init__(self, failures):
        self.failures = list(failures)
        self.calls = []
        self.chromium = self

    def launch_persistent_context(self, **kwargs):
        self.calls.append(kwargs)
        if self.failures:
            raise self.failures.pop(0)
        return _FakeContext()


BUSY = RuntimeError("Target page, context or browser has been closed ... exitCode=21, signal=null")


class LaunchPersistentTests(unittest.TestCase):
    def test_busy_profile_is_waited_for_then_launched_with_stealth(self):
        pw, waits = _FakePlaywright([BUSY, BUSY]), []
        context = bs.launch_persistent(pw, "prof", headless=False, disabled_features=("X",),
                                       delays=(1, 2, 3), sleep=waits.append)
        self.assertEqual(waits, [1, 2])
        self.assertEqual(context.scripts, [bs.STEALTH_INIT_SCRIPT])
        kwargs = pw.calls[-1]
        self.assertEqual((kwargs["channel"], kwargs["chromium_sandbox"]), ("chrome", True))
        self.assertIn("--disable-features=Translate,X", kwargs["args"])

    def test_gives_up_after_the_schedule_without_killing_anything(self):
        pw, waits = _FakePlaywright([BUSY] * 3), []
        with self.assertRaises(bs.ProfileBusy):
            bs.launch_persistent(pw, "prof", headless=True, delays=(1, 2), sleep=waits.append)
        self.assertEqual(waits, [1, 2])

    def test_other_launch_errors_are_not_retried(self):
        pw, waits = _FakePlaywright([RuntimeError("Executable doesn't exist")]), []
        with self.assertRaisesRegex(RuntimeError, "Executable"):
            bs.launch_persistent(pw, "prof", headless=True, delays=(1,), sleep=waits.append)
        self.assertEqual(waits, [])


if __name__ == "__main__":
    unittest.main()
