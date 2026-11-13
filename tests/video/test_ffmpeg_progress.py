"""The shared ffmpeg progress runner parses `time=` lines and reports percent."""

import subprocess
import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tools" / "video"))

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

    def test_utf8_bytes_cp1252_cannot_map_do_not_crash_the_runner(self):
        # "Ángel.mp4" in UTF-8 starts 0xC3 0x81; 0x81 is undefined in cp1252.
        script = (
            "import sys\n"
            "sys.stderr.buffer.write('Input #0, Ángel.mp4\\n'.encode('utf-8'))\n"
            "sys.stderr.buffer.write(b'time=00:00:05.00\\n')\n"
        )
        seen = []
        # Pinned explicitly: under PYTHONUTF8=1 the locale codec is already UTF-8 and would hide the bug.
        with mock.patch.object(_ffmpeg_utils.subprocess, "Popen", wraps=subprocess.Popen) as popen:
            code = _ffmpeg_utils.run_ffmpeg_with_progress([sys.executable, "-c", script], 10.0, seen.append)
        self.assertEqual(popen.call_args.kwargs["encoding"], "utf-8")
        self.assertEqual(code, 0)
        self.assertEqual(seen, [50.0])

    def test_failing_callback_does_not_orphan_ffmpeg(self):
        script = "import sys, time\nsys.stderr.write('time=00:00:05.00\\n')\nsys.stderr.flush()\ntime.sleep(60)\n"
        spawned = []
        real_popen = subprocess.Popen

        def capture(*args, **kwargs):
            spawned.append(real_popen(*args, **kwargs))
            return spawned[-1]

        def boom(_percent):
            raise RuntimeError("callback failed")

        with mock.patch.object(_ffmpeg_utils.subprocess, "Popen", side_effect=capture):
            with self.assertRaises(RuntimeError):
                _ffmpeg_utils.run_ffmpeg_with_progress([sys.executable, "-c", script], 10.0, boom)
        self.assertIsNotNone(spawned[0].poll())


if __name__ == "__main__":
    unittest.main()
