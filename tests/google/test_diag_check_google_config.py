"""Guard: the Google config diagnostic exercises what the gmail_drive automation
uses, not a flow the repo no longer has (offline, no network)."""

import json
import re
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from tools.google import diag_check_google_config as diag

_GOOGLE_DIR = Path(diag.__file__).resolve().parent


def _built_services(path: Path) -> set:
    """Names passed to `build('<service>', ...)` in a source file."""
    return set(re.findall(r"""\bbuild\(\s*['"](\w+)['"]""", path.read_text(encoding="utf-8")))


class DiagTargetsTheLiveFlowTests(unittest.TestCase):
    def test_it_only_builds_services_the_automation_builds(self) -> None:
        diag_services = _built_services(_GOOGLE_DIR / "diag_check_google_config.py")
        live_services = _built_services(_GOOGLE_DIR / "gmail_drive_automation.py")
        self.assertTrue(diag_services)
        self.assertLessEqual(diag_services, live_services)

    def test_credential_and_token_files_come_from_the_config(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            auth = {"credentials_file": "my_creds.json", "token_file": "my_token.json"}
            (folder / "config_gmail_drive.json.sample").write_text(json.dumps({"auth": auth}), encoding="utf-8")
            with mock.patch.object(diag, "_GOOGLE_DIR", folder):
                self.assertEqual(diag._auth_files(), (folder / "my_creds.json", folder / "my_token.json"))

    def test_the_shipped_sample_names_the_files_the_diagnostic_checks(self) -> None:
        sample = json.loads((_GOOGLE_DIR / "config_gmail_drive.json.sample").read_text(encoding="utf-8"))
        if (_GOOGLE_DIR / "config.json").exists():
            self.skipTest("a live config.json overrides the sample")
        creds, token = diag._auth_files()
        self.assertEqual((creds.name, token.name), (sample["auth"]["credentials_file"], sample["auth"]["token_file"]))


if __name__ == "__main__":
    unittest.main()
