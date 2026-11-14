"""The weekly Photos album job's preflight: a real run starts with its own dry run and stops on a mismatch (fakes only)."""

import json
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest import mock

from tools.google import photos_weekly_album as pwa
from tools.google import photos_weekly_album_job as job

SAMPLE = Path(pwa.__file__).with_name("photos_weekly_album.json.sample")
MISMATCH = ("saw 5 selectable tiles in the week but 4 were selected (differs on 2026-10-03: selected 2, tiles 3); "
            "the sweep missed items, so screenshots may remain")


class _Notifier:
    sent: list = []

    def __init__(self, env):
        self.env = env

    def send(self, text):
        self.sent.append(text)
        return True


def _config(tmp: Path) -> Path:
    raw = json.loads(SAMPLE.read_text(encoding="utf-8"))
    raw.update(profile_dir=str(tmp / "profile"), log_file="")
    path = tmp / "config.json"
    path.write_text(json.dumps(raw), encoding="utf-8")
    return path


class _Web:
    """A Photos client that has no create_album, share or mail method: a real pass reaching one fails the run."""

    cleared = 0
    dropped = 0

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
        raise pwa.FlowError(MISMATCH)

    def clear_selection(self):
        type(self).cleared += 1

    def drop_selection(self):
        type(self).dropped += 1


class PreflightTests(unittest.TestCase):
    def _execute(self, fake_run, dry_run=False):
        _Notifier.sent = []
        with tempfile.TemporaryDirectory() as tmp, mock.patch.object(job, "FleetNotifier", _Notifier), \
                self.assertLogs("photos_weekly_album_job") as logs:
            code = job.execute(_config(Path(tmp)), date(2026, 10, 10), {}, locked=lambda: False, run=fake_run,
                               dry_run=dry_run)
        return code, "\n".join(logs.output)

    def test_clean_preflight_is_followed_by_the_real_run(self):
        calls = []

        def fake_run(cfg, start, end, title, dry_run, send=False, record=None):
            calls.append(dry_run)
            if dry_run:
                record.update(outcome="dry-run", selected=3, per_day={"2026-10-03": 3}, tiles_per_day={"2026-10-03": 3})
            else:
                record.update(outcome="created", url="https://photos.example/album", mail=None)
            return 0

        code, logs = self._execute(fake_run)
        self.assertEqual((code, calls), (0, [True, False]))
        self.assertEqual(len(_Notifier.sent), 1)
        self.assertIn("Weekly album done", _Notifier.sent[0])
        self.assertIn("preflight", logs)
        self.assertIn("2026-10-03", logs)

    def test_a_count_mismatch_in_the_preflight_stops_before_the_real_run_and_alerts(self):
        calls = []

        def fake_run(cfg, start, end, title, dry_run, send=False, record=None):
            calls.append(dry_run)
            record.update(outcome="error", failure="error", error=MISMATCH)
            return 1

        code, _ = self._execute(fake_run)
        self.assertEqual((code, calls), (job.EXIT["preflight-failed"], [True]))
        self.assertNotIn(code, (0, 1, 2, 3, 4, 5))
        self.assertEqual(len(_Notifier.sent), 1)
        alert = _Notifier.sent[0]
        self.assertIn("preflight", alert)
        self.assertIn("2026-10-03: selected 2, tiles 3", alert)
        self.assertIn("Nothing was created", alert)

    def test_a_preflight_that_ends_without_a_dry_run_outcome_does_not_count_as_clean(self):
        calls = []

        def fake_run(cfg, start, end, title, dry_run, send=False, record=None):
            calls.append(dry_run)
            record.update(outcome="exists")
            return 0

        code, _ = self._execute(fake_run)
        self.assertEqual((code, calls), (job.EXIT["preflight-failed"], [True]))
        self.assertEqual(len(_Notifier.sent), 1)

    def test_a_preflight_infrastructure_failure_keeps_its_own_exit_code_and_message(self):
        calls = []

        def fake_run(cfg, start, end, title, dry_run, send=False, record=None):
            calls.append(dry_run)
            record.update(outcome="not-signed-in", failure="not-signed-in")
            return 3

        code, _ = self._execute(fake_run)
        self.assertEqual((code, calls), (3, [True]))
        self.assertEqual(len(_Notifier.sent), 1)
        self.assertIn("signed out", _Notifier.sent[0])

    def test_the_jobs_own_dry_run_is_a_single_pass(self):
        calls = []

        def fake_run(cfg, start, end, title, dry_run, send=False, record=None):
            calls.append(dry_run)
            record.update(outcome="dry-run")
            return 0

        code, _ = self._execute(fake_run, dry_run=True)
        self.assertEqual((code, calls, _Notifier.sent), (0, [True], []))

    def test_through_main_a_mismatch_creates_nothing_clears_the_selection_and_logs_the_preflight_only(self):
        _Notifier.sent = []
        _Web.cleared = _Web.dropped = 0
        with tempfile.TemporaryDirectory() as tmp:
            runtime = Path(tmp) / "runtime"
            with mock.patch.dict("os.environ", {"AUTOMATION_RUNTIME_DIR": str(runtime)}), \
                    mock.patch.object(job, "desktop_locked", return_value=False), \
                    mock.patch.object(job, "FleetNotifier", _Notifier), \
                    mock.patch.object(pwa, "PhotosWeb", _Web), \
                    self.assertLogs("photos_weekly_album_job"):
                code = job.main(["--config", str(_config(Path(tmp)))])
            lines = (runtime / "photos_weekly_album.log").read_text(encoding="utf-8").splitlines()
        self.assertEqual(code, job.EXIT["preflight-failed"])
        self.assertEqual(_Web.dropped, 1)
        self.assertEqual(len(_Notifier.sent), 1)
        self.assertEqual(len(lines), 1)
        record = json.loads(lines[0])
        self.assertEqual((record["dry_run"], record["preflight"]), (True, True))


if __name__ == "__main__":
    unittest.main()
