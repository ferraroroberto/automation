"""The weekly Photos album job's --dry-run, through the real entry point the .bat calls (no browser, no network, no mail)."""

import json
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest import mock

from tools.google import photos_weekly_album as pwa
from tools.google import photos_weekly_album_job as job

SAMPLE = Path(pwa.__file__).with_name("photos_weekly_album.json.sample")


class _Notifier:
    sent: list = []

    def __init__(self, env):
        self.env = env

    def send(self, text):
        self.sent.append(text)
        return True


class _DryRunWeb:
    """Only what a dry run may touch: no create_album, share or mail method exists, so reaching one fails the run."""

    cleared = 0

    def __init__(self, profile_dir):
        pass

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        pass

    def require_signed_in(self):
        pass

    def find_album(self, title):
        return None

    def collect(self, start, end):
        return pwa.Selection(per_day={start: 3}, tiles_per_day={start: 3}, selected=3)

    def clear_selection(self):
        type(self).cleared += 1


def _repo_files() -> set:
    return {p for p in Path(job.REPO_ROOT, "tools", "google").rglob("*")
            if p.is_file() and "__pycache__" not in p.parts}


def _config(tmp: Path) -> Path:
    raw = json.loads(SAMPLE.read_text(encoding="utf-8"))
    raw.update(profile_dir=str(tmp / "profile"), log_file="")  # "" takes the default log location
    path = tmp / "config.json"
    path.write_text(json.dumps(raw), encoding="utf-8")
    return path


class JobDryRunEntryPointTests(unittest.TestCase):
    def test_dry_run_through_main_writes_nothing_in_the_repo_and_sends_nothing(self):
        _Notifier.sent = []
        _DryRunWeb.cleared = 0
        before = _repo_files()
        with tempfile.TemporaryDirectory() as tmp:
            runtime = Path(tmp) / "runtime"
            with mock.patch.dict("os.environ", {"AUTOMATION_RUNTIME_DIR": str(runtime)}), \
                    mock.patch.object(job, "desktop_locked", return_value=False), \
                    mock.patch.object(job, "FleetNotifier", _Notifier), \
                    mock.patch.object(pwa, "PhotosWeb", _DryRunWeb), \
                    self.assertLogs("photos_weekly_album_job"):
                code = job.main(["--dry-run", "--config", str(_config(Path(tmp)))])
            log = runtime / "photos_weekly_album.log"
            log_lines = log.read_text(encoding="utf-8").splitlines() if log.exists() else []
        self.assertEqual(code, 0)
        self.assertEqual(_Notifier.sent, [])
        self.assertEqual(_DryRunWeb.cleared, 1)
        self.assertEqual(len(log_lines), 1)
        record = json.loads(log_lines[0])
        self.assertEqual((record["outcome"], record["dry_run"], record["selected"]), ("dry-run", True, 3))
        self.assertEqual(_repo_files() - before, set())

    def test_dry_run_flag_reaches_the_album_run_and_skips_the_notice_on_failure(self):
        calls = []

        def fake_run(cfg, start, end, title, dry_run, send=False, record=None):
            calls.append(dry_run)
            record.update(outcome="error", failure="window-hidden")
            return 4

        _Notifier.sent = []
        with tempfile.TemporaryDirectory() as tmp, mock.patch.object(job, "FleetNotifier", _Notifier), \
                self.assertLogs("photos_weekly_album_job"):
            code = job.execute(_config(Path(tmp)), date(2026, 10, 10), {}, locked=lambda: False,
                               run=fake_run, dry_run=True)
        self.assertEqual((code, calls, _Notifier.sent), (4, [True], []))


if __name__ == "__main__":
    unittest.main()
