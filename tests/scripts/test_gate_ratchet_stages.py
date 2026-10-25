"""Every ratchet `.fleet.toml` enables must run as a gate stage that fails on exit 3.

`scripts/verify-before-ship.ps1` is the one ship gate. A ratchet that is declared and
vendored but never invoked there, or invoked without the exit-3 branch, would let
"could not measure" read as a pass, so the wiring is pinned from the config side.
"""

import re
import tomllib
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
GATE = (REPO_ROOT / "scripts" / "verify-before-ship.ps1").read_text(encoding="utf-8")
FLEET = tomllib.loads((REPO_ROOT / ".fleet.toml").read_text(encoding="utf-8"))

ENABLED = sorted(
    name for name, table in FLEET.items()
    if name.endswith("_ratchet") and isinstance(table, dict) and table.get("enabled") is True
)


class GateRatchetStages(unittest.TestCase):
    def test_enabled_ratchets_are_found(self) -> None:
        self.assertTrue(ENABLED, "no enabled [*_ratchet] table found in .fleet.toml")

    def test_each_enabled_ratchet_is_vendored(self) -> None:
        for name in ENABLED:
            with self.subTest(ratchet=name):
                entry = FLEET.get("vendored", {}).get(name)
                self.assertIsNotNone(entry, f"{name} is enabled but not in [vendored]")
                self.assertEqual(entry["dest"], f"scripts/{name}.py")
                self.assertRegex(entry["sha"], r"^[0-9a-f]{40}$")
                self.assertTrue((REPO_ROOT / entry["dest"]).is_file())

    def test_each_enabled_ratchet_has_a_baseline_file(self) -> None:
        for name in ENABLED:
            with self.subTest(ratchet=name):
                baseline = FLEET[name].get("baseline")
                self.assertTrue(baseline, f"[{name}] declares no baseline")
                self.assertTrue((REPO_ROOT / baseline).is_file())

    def test_each_enabled_ratchet_is_a_gate_stage_that_fails_on_exit_3(self) -> None:
        for name in ENABLED:
            with self.subTest(ratchet=name):
                stage = re.search(
                    rf'& \$python "scripts/{name}\.py"\s*\n'
                    r'\s*if \(\$LASTEXITCODE -eq 3\) \{ Fail "[^"]*NOT CONFIRMED[^"]*" \}\s*\n'
                    r'\s*if \(\$LASTEXITCODE -ne 0\) \{ Fail ',
                    GATE,
                )
                self.assertIsNotNone(stage, f"{name} has no gate stage with an exit-3 NOT CONFIRMED branch")


if __name__ == "__main__":
    unittest.main()
