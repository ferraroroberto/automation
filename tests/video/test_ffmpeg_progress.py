"""The shared ffmpeg progress runner parses `time=` lines and reports percent."""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "video"))

import _ffmpeg_utils  # noqa: E402


def _fake_ffmpeg(stderr_lines, exit_code=0):
    """A command that prints `stderr_lines` to stderr and exits with `exit_code`."""
    script = "import sys\n" + "".join(f"sys.stderr.write({line!r} + '\\n')\n" for line in stderr_lines)
    script += f"sys.exit({exit_code})"
    return [sys.executable, "-c", script]


class ParseFfmpegTimeTests(unittest.TestCase):
    def test_hours_minutes_seconds(self):
        self.assertEqual(_ffmpeg_utils.parse_ffmpeg_time("frame=1 time=01:02:03.50 bitrate=1"), 3723.5)

    def test_minutes_seconds(self):
        self.assertEqual(_ffmpeg_utils.parse_ffmpeg_time("time=02:03.25"), 123.25)

    def test_seconds_only(self):
        self.assertEqual(_ffmpeg_utils.parse_ffmpeg_time("time=7.50"), 7.5)

    def test_line_without_time_is_none(self):
        self.assertIsNone(_ffmpeg_utils.parse_ffmpeg_time("Press [q] to stop"))


class RunFfmpegWithProgressTests(unittest.TestCase):
    def test_reports_percent_and_returns_exit_code(self):
        seen = []
        code = _ffmpeg_utils.run_ffmpeg_with_progress(
            _fake_ffmpeg(["time=00:00:05.00", "time=00:00:10.00"]), 10.0, seen.append
        )
        self.assertEqual(code, 0)
        self.assertEqual(seen, [50.0, 100.0])

    def test_max_percent_caps_the_report(self):
        seen = []
        _ffmpeg_utils.run_ffmpeg_with_progress(
            _fake_ffmpeg(["time=00:00:10.00"]), 10.0, seen.append, max_percent=99.0
        )
        self.assertEqual(seen, [99.0])

    def test_unknown_duration_falls_back_to_rough_estimate(self):
        seen = []
        _ffmpeg_utils.run_ffmpeg_with_progress(_fake_ffmpeg(["time=00:00:10.00"]), None, seen.append)
        self.assertEqual(seen, [20.0])

    def test_nonzero_exit_code_is_returned(self):
        code = _ffmpeg_utils.run_ffmpeg_with_progress(_fake_ffmpeg([], exit_code=3), 10.0)
        self.assertEqual(code, 3)


if __name__ == "__main__":
    unittest.main()
