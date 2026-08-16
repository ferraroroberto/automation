#!/usr/bin/env python3
r"""Guard: gpu_recovery must read its console tools the same in any locale.

The bug these tests lock down (issue #108, the same class as #106): a
`subprocess.run(..., text=True)` spawn decodes the child's output with the
*parent's* ambient text encoding, not the child's console code page.

Measured on this machine (Windows 11, OEM 850, ACP 1252) against a child
emitting the cp850 bytes ``Configuraci\xa2n``:

* `PYTHONUTF8` unset - `text=True` decodes as cp1252 and silently yields the
  wrong character (``¢`` for ``ó``). Exit 0, corrupt string.
* `PYTHONUTF8=1` - `UnicodeDecodeError` is raised inside subprocess's reader
  thread, which prints a traceback and dies. `CompletedProcess` comes back
  with `returncode == 0` and `stdout is None`, so the old
  `result.stdout.strip()` raised `AttributeError` - not caught by
  `except (subprocess.SubprocessError, OSError)`, so all three pre-fix call
  sites died with an unhandled traceback.

Either way the tri-state matters: `nvidia_smi()` must never fold "could not
read the GPU" into `None`-meaning-dead, on the one alarm this tool exists to
raise.

Run from the repo root:

    & .\.venv\Scripts\python.exe -m unittest discover -s video -p "test_*.py"
"""

from __future__ import annotations

import ast
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

VIDEO_DIR = Path(__file__).resolve().parent
REPO_ROOT = VIDEO_DIR.parent
sys.path.insert(0, str(VIDEO_DIR))

import gpu_recovery  # noqa: E402

sys.path.insert(0, str(REPO_ROOT / "system"))
from _lib import console_output  # noqa: E402

# The cp850 encoding of "ó" (0xa2) - deliberately *not* valid UTF-8, which is
# what a localized Windows console tool's output looks like on this machine.
CP850_TEXT_BYTES = b"NVIDIA GeForce RTX 5060 Ti, Configuraci\xa2n, 16311 MiB\r\n"

# A stand-in child that writes those bytes straight to stdout and exits 0,
# i.e. behaves exactly like a native console tool with a localized message.
STAND_IN_SOURCE = (
    "import sys\n"
    "sys.stdout.buffer.write(%r)\n"
    "sys.exit(0)\n" % CP850_TEXT_BYTES
)


class DecoderPromotionTests(unittest.TestCase):
    """#106's decoder moved to system/_lib - it must not have forked."""

    def test_gpu_recovery_uses_the_shared_decoder(self) -> None:
        self.assertIs(gpu_recovery.decode_console_bytes, console_output.decode_console_bytes)

    def test_netsh_still_exposes_the_same_decoder_object(self) -> None:
        sys.path.insert(0, str(REPO_ROOT / "system" / "wifi"))
        import _netsh  # noqa: PLC0415

        self.assertIs(_netsh.decode_console_bytes, console_output.decode_console_bytes)
        self.assertIs(_netsh.NO_WINDOW, console_output.NO_WINDOW)
        self.assertEqual(_netsh.NETSH_ENCODINGS, console_output.CONSOLE_ENCODINGS)

    def test_non_utf8_console_bytes_survive_the_decode(self) -> None:
        text, encoding = console_output.decode_console_bytes(CP850_TEXT_BYTES)
        self.assertIn("NVIDIA GeForce RTX 5060 Ti", text)
        self.assertIn("16311 MiB", text)
        self.assertNotEqual(encoding, "utf-8")


class NoInheritedDecodingTests(unittest.TestCase):
    """Acceptance criterion 1, enforced against the source."""

    @staticmethod
    def _spawn_calls(source: str) -> list[str]:
        """Return the argument text of every `subprocess.run(...)` in a source."""
        calls = []
        for chunk in source.split("subprocess.run(")[1:]:
            depth, end = 1, 0
            for end, char in enumerate(chunk):
                depth += (char == "(") - (char == ")")
                if depth == 0:
                    break
            calls.append(chunk[:end])
        return calls

    def test_gpu_recovery_spawns_once_through_the_shared_runner(self) -> None:
        source = (VIDEO_DIR / "gpu_recovery.py").read_text(encoding="utf-8")
        self.assertEqual(
            len(self._spawn_calls(source)), 1, "every console tool must go through _run_console"
        )

    def test_the_spawn_never_decodes_with_the_parent_locale(self) -> None:
        source = (VIDEO_DIR / "gpu_recovery.py").read_text(encoding="utf-8")
        for call in self._spawn_calls(source):
            self.assertNotIn("text=", call, "the spawn decodes with the parent's locale")
            self.assertNotIn("encoding=", call, "the spawn pins a decode instead of capturing bytes")

    def test_the_spawn_captures_bytes_and_suppresses_the_console_window(self) -> None:
        source = (VIDEO_DIR / "gpu_recovery.py").read_text(encoding="utf-8")
        for call in self._spawn_calls(source):
            self.assertIn("capture_output=True", call)
            self.assertIn("creationflags=_NO_WINDOW", call)


class TriStateTests(unittest.TestCase):
    """"Could not read the GPU" must never be reported as "the GPU is gone"."""

    def test_unrunnable_smi_is_unknown_not_dead(self) -> None:
        original = gpu_recovery._smi_command
        gpu_recovery._smi_command = lambda: ["nvidia-smi-does-not-exist-here"]
        try:
            reading = gpu_recovery.nvidia_smi()
        finally:
            gpu_recovery._smi_command = original
        self.assertIsNone(reading.summary)
        self.assertFalse(reading.known, "a failed spawn must report unknown, not 'not live'")

    def test_zero_exit_with_unreadable_output_is_unknown(self) -> None:
        run = gpu_recovery.ConsoleRun(
            ran=True, returncode=0, stdout="", stderr="", encoding="utf-8"
        )
        original = gpu_recovery._run_console
        gpu_recovery._run_console = lambda command, *, timeout, label: run
        try:
            with self.assertLogs(gpu_recovery.logger, level="ERROR"):
                reading = gpu_recovery.nvidia_smi()
        finally:
            gpu_recovery._run_console = original
        self.assertIsNone(reading.summary)
        self.assertFalse(reading.known)

    def test_nonzero_exit_is_a_real_negative_answer(self) -> None:
        run = gpu_recovery.ConsoleRun(
            ran=True, returncode=9, stdout="", stderr="no devices", encoding="utf-8"
        )
        original = gpu_recovery._run_console
        gpu_recovery._run_console = lambda command, *, timeout, label: run
        try:
            reading = gpu_recovery.nvidia_smi()
        finally:
            gpu_recovery._run_console = original
        self.assertIsNone(reading.summary)
        self.assertTrue(reading.known, "nvidia-smi refusing IS an established 'not live'")

    def test_failed_adapter_query_is_unknown_not_absent(self) -> None:
        run = gpu_recovery.ConsoleRun(
            ran=False, returncode=-1, stdout="", stderr="", encoding=None
        )
        original = gpu_recovery._run_console
        gpu_recovery._run_console = lambda command, *, timeout, label: run
        try:
            query = gpu_recovery.find_nvidia_gpus()
        finally:
            gpu_recovery._run_console = original
        self.assertEqual(query.gpus, [])
        self.assertFalse(query.known, "an unreadable query must not mean 'card absent'")

    def test_empty_pipeline_is_a_real_absent_answer(self) -> None:
        run = gpu_recovery.ConsoleRun(
            ran=True, returncode=0, stdout="", stderr="", encoding="utf-8"
        )
        original = gpu_recovery._run_console
        gpu_recovery._run_console = lambda command, *, timeout, label: run
        try:
            query = gpu_recovery.find_nvidia_gpus()
        finally:
            gpu_recovery._run_console = original
        self.assertEqual(query.gpus, [])
        self.assertTrue(query.known, "ConvertTo-Json prints nothing for an empty pipeline")


def _read_smi_in_child(utf8_mode: str, stand_in: Path) -> subprocess.CompletedProcess:
    """Call `nvidia_smi()` in a child with PYTHONUTF8 pinned, via a stand-in tool.

    The child hands its answer back through `stdout.buffer` as explicit UTF-8
    rather than `print()`, because `print()` would re-encode through the
    child's *own* locale - the very variable under test - and could mask a
    difference the nvidia-smi decode had actually introduced.

    The stand-in is a plain Python child writing fixed bytes, so nothing here
    depends on the machine having a GPU: a non-zero exit is a real failure,
    never an environment to skip over.
    """
    code = (
        "import sys; sys.path.insert(0, r'%s');\n"
        "import gpu_recovery;\n"
        "gpu_recovery._smi_command = lambda: [sys.executable, r'%s'];\n"
        "reading = gpu_recovery.nvidia_smi();\n"
        "sys.stdout.buffer.write(repr((reading.summary, reading.known)).encode('utf-8'))\n"
        % (str(VIDEO_DIR), str(stand_in))
    )
    env = dict(os.environ)
    env["PYTHONUTF8"] = utf8_mode
    return subprocess.run(
        [sys.executable, "-c", code],
        capture_output=True,
        env=env,
        cwd=str(VIDEO_DIR),
        creationflags=gpu_recovery._NO_WINDOW,
    )


class LiveReadUnderUtf8ModeTests(unittest.TestCase):
    """The regression: a localized tool reads the same whatever the parent's mode.

    Before the fix these two runs disagreed - with `PYTHONUTF8=1` the reader
    thread's UnicodeDecodeError emptied stdout and `nvidia_smi()` returned
    `None`, i.e. "the GPU isn't live", while with it off the same bytes came
    back mangled through cp1252.
    """

    def setUp(self) -> None:
        # Outside the repo: a stray .py in video/ would land in compileall.
        tmp = tempfile.mkdtemp(prefix="gpu_recovery_test_")
        self.addCleanup(shutil.rmtree, tmp, True)
        self.stand_in = Path(tmp) / "stand_in_smi.py"
        self.stand_in.write_text(STAND_IN_SOURCE, encoding="utf-8")

    def _read(self, utf8_mode: str) -> str:
        proc = _read_smi_in_child(utf8_mode, self.stand_in)
        self.assertEqual(
            proc.returncode,
            0,
            f"PYTHONUTF8={utf8_mode} child failed:\n"
            f"{proc.stderr.decode('utf-8', 'replace')}",
        )
        return proc.stdout.decode("utf-8")

    def test_summary_is_read_under_utf8_mode(self) -> None:
        summary, known = ast.literal_eval(self._read("1"))
        self.assertTrue(known, "the reading must be established, not unknown")
        self.assertIsNotNone(summary, "PYTHONUTF8=1 emptied the output - #108 is back")
        self.assertIn("NVIDIA GeForce RTX 5060 Ti", summary)
        self.assertIn("16311 MiB", summary)

    def test_reading_matches_with_utf8_mode_on_and_off(self) -> None:
        on = self._read("1")
        off = self._read("0")
        self.assertEqual(on, off, "the GPU reading still depends on the parent's locale")


if __name__ == "__main__":
    unittest.main()
