r"""Guard: machine-local config files must stay out of version control.

Several tools in this repo read a local config file that is generated on, or
tailored to, one machine. The repo convention is that only a `.sample` /
`.example` template is tracked and the live file is gitignored. Two failure
modes have bitten this repo before:

1. A live config was committed before the ignore rule existed. `.gitignore`
   does not apply retroactively to already-tracked paths, so the rule looks
   present while the file stays tracked forever.
2. A tool's default output path collided with its own tracked config file, so
   a plain run rewrote a tracked file with locally-generated content.

This test locks both down. Run it directly:

    & .\.venv\Scripts\python.exe system\test_local_config_hygiene.py
"""

from __future__ import annotations

import json
import re
import subprocess
import sys
import unittest
from pathlib import Path
from typing import List, Set

REPO_ROOT = Path(__file__).resolve().parent.parent

NO_WINDOW = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0

# Paths that must never appear in `git ls-files`. Each is a machine-local file
# whose tracked counterpart is a .sample/.example template.
MUST_NOT_BE_TRACKED: List[str] = [
    "google/config_gmail_drive.json",
    "google/config_weekly_photo.json",
    "linkedin/open_profiles/linkedin_open.json",
    "system/textexpander/text_expander_config.json",
    "system/wifi/wifi_connect.xml",
    "system/wifi/wifi_export.json",
]


def tracked_files() -> Set[str]:
    """Return every path tracked in git, using forward slashes."""
    result = subprocess.run(
        ["git", "ls-files"],
        cwd=REPO_ROOT,
        capture_output=True,
        text=True,
        check=True,
        creationflags=NO_WINDOW,
    )
    return {line.strip().replace("\\", "/") for line in result.stdout.splitlines() if line.strip()}


class LocalConfigHygieneTests(unittest.TestCase):
    def test_machine_local_config_files_are_not_tracked(self) -> None:
        tracked = tracked_files()
        offenders = sorted(path for path in MUST_NOT_BE_TRACKED if path in tracked)
        self.assertEqual(
            offenders,
            [],
            "These machine-local files are tracked in git; untrack them with "
            "`git rm --cached <path>` and keep only the .sample template: "
            f"{offenders}",
        )

    def test_wifi_export_default_does_not_overwrite_tracked_config(self) -> None:
        """The exporter's default output must not be its own tracked config file."""
        config_path = REPO_ROOT / "system" / "wifi" / "wifi_passwords.json"
        if not config_path.exists():
            self.skipTest("system/wifi/wifi_passwords.json is absent")

        config = json.loads(config_path.read_text(encoding="utf-8"))
        default_output = config.get("default_output")
        self.assertIsNotNone(default_output, "wifi_passwords.json has no default_output key")
        self.assertNotEqual(
            default_output,
            config_path.name,
            "default_output points at the tracked config file itself, so a default "
            "run would overwrite it with locally-generated content",
        )
        self.assertNotIn(
            f"system/wifi/{default_output}",
            tracked_files(),
            f"default_output '{default_output}' is a tracked path; a default run would "
            "rewrite a tracked file",
        )

    def test_wifi_export_fallback_default_does_not_overwrite_tracked_config(self) -> None:
        """load_config()'s in-code fallback (used when wifi_passwords.json can't be
        read) must not default to the tracked config file's own name either."""
        wifi_dir = REPO_ROOT / "system" / "wifi"
        source = (wifi_dir / "wifi_passwords.py").read_text(encoding="utf-8")
        match = re.search(
            r"#\s*Fallback default configuration\s*\n\s*return\s*\{\s*\n\s*\"default_output\"\s*:\s*\"([^\"]+)\"",
            source,
        )
        self.assertIsNotNone(
            match,
            "Could not locate load_config()'s fallback default_output literal — "
            "update this test if load_config() was restructured",
        )
        self.assertNotEqual(
            match.group(1),
            "wifi_passwords.json",
            "load_config()'s fallback default_output points at the tracked config "
            "file itself, so a run with a missing/unreadable config would overwrite it",
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)
