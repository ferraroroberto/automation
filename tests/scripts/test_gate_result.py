"""The pre-ship gate ends every run with a `GATE-RESULT: PASS|FAIL` stdout line.

A pipeline reports its last command's status, so `verify-before-ship.ps1 | tail` hands
the caller `tail`'s 0 and a red gate reads green. The verdict line is the part a pipe
cannot rewrite. The helper block in the gate (between the `gate-verdict` markers) is run
for real, over a captured stdout (the same anonymous pipe `|` creates), for a pass, a
failed stage and an early terminating error; the real script is run for the paths that
need no stage to execute.
"""

import os
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.append(str(REPO_ROOT))

from no_window import NO_WINDOW  # noqa: E402

GATE = REPO_ROOT / "scripts" / "verify-before-ship.ps1"
BOX_LOCK = REPO_ROOT / "scripts" / "box_lock_gate.ps1"
POWERSHELL = Path(os.environ.get("SystemRoot", "")) / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe"

# The gate is a Windows PowerShell script; the suite's other tests run anywhere.
windows_only = unittest.skipUnless(sys.platform == "win32" and POWERSHELL.is_file(), "needs Windows PowerShell")


def _gate_text() -> str:
    return GATE.read_text(encoding="utf-8").replace("\r\n", "\n")


def _verdict_block() -> str:
    match = re.search(r"# gate-verdict:begin\n(.*?)# gate-verdict:end", _gate_text(), re.DOTALL)
    assert match, "the gate lost its gate-verdict block"
    return match.group(1)


def _run(args: list[str], env: dict[str, str] | None = None) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [str(POWERSHELL), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", *args],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        stdin=subprocess.DEVNULL, creationflags=NO_WINDOW, timeout=120, env=env,
    )


def _run_scenario(scenario: str) -> subprocess.CompletedProcess[str]:
    # The gate defines $sw before the block; Fail reads it.
    script = f'$ErrorActionPreference = "Stop"\n$sw = [System.Diagnostics.Stopwatch]::StartNew()\n{_verdict_block()}\n{scenario}\n'
    return _run(["-Command", script])


def _last_line(run: subprocess.CompletedProcess[str]) -> str:
    lines = [line.rstrip() for line in run.stdout.splitlines() if line.strip()]
    assert lines, f"nothing on stdout; stderr={run.stderr!r}"
    return lines[-1]


def _gate_copy(root: Path) -> Path:
    scripts = root / "scripts"
    scripts.mkdir()
    shutil.copy2(GATE, scripts / GATE.name)
    shutil.copy2(BOX_LOCK, scripts / BOX_LOCK.name)
    return scripts / GATE.name


def _isolated_env(root: Path) -> dict[str, str]:
    # An empty profile has no box_lock.py, so the gate runs unlocked and never reaches the real helper.
    return {**os.environ, "USERPROFILE": str(root)}


@windows_only
class GateResultLine(unittest.TestCase):
    def test_pass_ends_with_pass_and_exits_zero(self) -> None:
        run = _run_scenario("Complete-Gate 0")
        self.assertEqual(run.returncode, 0, run.stdout)
        self.assertEqual(_last_line(run), "GATE-RESULT: PASS")

    def test_failed_stage_ends_with_fail_and_exits_one(self) -> None:
        run = _run_scenario('cmd /c exit 7\nif ($LASTEXITCODE -ne 0) { Fail "stage bad failed." }\nComplete-Gate 0')
        self.assertEqual(run.returncode, 1, run.stdout)
        self.assertEqual(_last_line(run), "GATE-RESULT: FAIL")
        self.assertIn("[X] stage bad failed.", run.stdout)

    def test_early_terminating_error_reports_on_stdout_then_fails(self) -> None:
        run = _run_scenario('throw "boom before any stage"')
        self.assertEqual(run.returncode, 1, run.stdout)
        self.assertEqual(_last_line(run), "GATE-RESULT: FAIL")
        self.assertIn("boom before any stage", run.stdout, "the error must reach stdout, not only stderr")

    def test_missing_venv_ends_with_fail_and_exits_one(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            run = _run(["-File", str(_gate_copy(root))], env=_isolated_env(root))
        self.assertEqual(run.returncode, 1, run.stdout)
        self.assertEqual(_last_line(run), "GATE-RESULT: FAIL")
        self.assertIn(".venv missing", run.stdout)

    def test_unrunnable_interpreter_is_reported_as_fail_not_a_silent_run(self) -> None:
        """A terminating error in the real script: the interpreter path exists but is not an executable."""
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            gate = _gate_copy(root)
            interpreter = root / ".venv" / "Scripts" / "python.exe"
            interpreter.parent.mkdir(parents=True)
            interpreter.write_text("not an executable", encoding="utf-8")
            run = _run(["-File", str(gate)], env=_isolated_env(root))
        self.assertEqual(run.returncode, 1, run.stdout)
        self.assertEqual(_last_line(run), "GATE-RESULT: FAIL")


class GateScriptShape(unittest.TestCase):
    def test_every_exit_goes_through_the_verdict(self) -> None:
        """A new `exit` that skips `Complete-Gate` would end a run with no verdict."""
        text = _gate_text()
        outside = text.replace(_verdict_block(), "")
        self.assertIsNone(re.search(r"^\s*exit\s+\d", outside, re.MULTILINE), "exit through Complete-Gate, not directly")
        self.assertTrue(text.rstrip().endswith("Complete-Gate 0"), "the pass path must end on its verdict")

    def test_verdict_block_is_ascii_for_windows_powershell_5_1(self) -> None:
        """Windows PowerShell 5.1 reads a BOM-less file as ANSI, so the block stays ASCII."""
        _verdict_block().encode("ascii")


if __name__ == "__main__":
    unittest.main()
