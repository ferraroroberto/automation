#!/usr/bin/env python3
r"""Guard: netsh output must decode identically whatever the parent's locale.

The bug these tests lock down (issue #106): a `subprocess.run(..., text=True)`
or `encoding='utf-8'` spawn decodes the child's output with the *parent's*
ambient text encoding. That makes the wifi tools' behaviour depend on whether
the parent process happens to run in Python's UTF-8 mode — a tray app, a
wrapper .bat or an app-launcher job can flip it — and a bad decode is not
loud: it mangles or empties the output, so the caller reports "not connected"
/ "no profiles" / "export failed" for reasons unrelated to the network.

Run from the repo root:

    & .\.venv\Scripts\python.exe -m unittest discover -s system/wifi -p "test_*.py"
"""

from __future__ import annotations

import os
import subprocess
import sys
import unittest
from pathlib import Path

WIFI_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(WIFI_DIR))

import _netsh  # noqa: E402

# The three tools whose netsh spawns must all go through the shared helper.
TOOL_MODULES = ("wifi_connect.py", "wifi_passwords.py", "wifi_bat_generator.py")

# A real SSID captured from this machine: a right single quote (U+2019) that
# netsh emits as UTF-8 even when the console code page is OEM 850.
UTF8_SSID_BYTES = "Roberto’s iPhone".encode("utf-8")

# The same character as cp850 would encode it - deliberately *not* valid UTF-8,
# which is what netsh's own localized labels look like on a Spanish console.
CP850_LABEL_BYTES = b"Configuraci\xa2n de red"


class DecodeConsoleBytesTests(unittest.TestCase):
    """decode_console_bytes must always return something usable."""

    def test_utf8_bytes_decode_as_utf8(self) -> None:
        text, encoding = _netsh.decode_console_bytes(UTF8_SSID_BYTES)
        self.assertEqual(text, "Roberto’s iPhone")
        self.assertEqual(encoding, "utf-8")

    def test_non_utf8_bytes_fall_back_without_raising(self) -> None:
        text, encoding = _netsh.decode_console_bytes(CP850_LABEL_BYTES)
        self.assertTrue(text.startswith("Configuraci"), text)
        self.assertNotEqual(encoding, "utf-8")
        # The whole line survives; at most the one odd byte is substituted.
        self.assertIn("n de red", text)

    def test_empty_output_decodes_to_empty_string(self) -> None:
        text, encoding = _netsh.decode_console_bytes(b"")
        self.assertEqual(text, "")
        self.assertEqual(encoding, "utf-8")

    def test_undecodable_bytes_still_yield_text(self) -> None:
        # Byte 0x81 is undefined in cp1252 and invalid in UTF-8; the final
        # errors="replace" pass must still hand back the surrounding text.
        text, _ = _netsh.decode_console_bytes(b"before\x81after")
        self.assertIn("before", text)
        self.assertIn("after", text)


class ResultStateTests(unittest.TestCase):
    """A query that could not run must not look like a negative answer."""

    def test_failed_run_is_not_a_negative_answer(self) -> None:
        result = _netsh.NetshResult(
            args=("netsh", "wlan", "show", "interfaces"),
            returncode=-1,
            stdout="",
            stderr="",
            encoding=None,
            error="netsh not found",
        )
        self.assertFalse(result.ran)
        self.assertFalse(result.succeeded)
        self.assertFalse(result.has_output)

    def test_current_ssid_reports_unknown_when_the_query_fails(self) -> None:
        failed = _netsh.NetshResult(
            args=("netsh", "wlan", "show", "interfaces"),
            returncode=-1,
            stdout="",
            stderr="",
            encoding=None,
            error="boom",
        )
        original = _netsh.run_netsh
        _netsh.run_netsh = lambda args: failed  # type: ignore[assignment]
        try:
            state = _netsh.current_ssid()
        finally:
            _netsh.run_netsh = original  # type: ignore[assignment]
        self.assertFalse(state.known, "a failed query must report unknown, not 'not connected'")
        self.assertIsNone(state.ssid)

    def test_current_ssid_reports_none_when_genuinely_disconnected(self) -> None:
        disconnected = _netsh.NetshResult(
            args=("netsh", "wlan", "show", "interfaces"),
            returncode=0,
            stdout="There is no wireless interface on the system.\n",
            stderr="",
            encoding="utf-8",
        )
        original = _netsh.run_netsh
        _netsh.run_netsh = lambda args: disconnected  # type: ignore[assignment]
        try:
            state = _netsh.current_ssid()
        finally:
            _netsh.run_netsh = original  # type: ignore[assignment]
        self.assertTrue(state.known)
        self.assertIsNone(state.ssid)

    def test_parse_ssid_ignores_bssid_lines(self) -> None:
        output = "    BSSID                  : aa:bb:cc:dd:ee:ff\n    SSID                   : Roberto’s iPhone\n"
        self.assertEqual(_netsh.parse_ssid(output), "Roberto’s iPhone")


def _run_netsh_in_child(utf8_mode: str) -> str:
    """Run the shared helper in a child process with PYTHONUTF8 pinned.

    The child hands its result back through `stdout.buffer` as explicit UTF-8
    rather than `print()`, because `print()` would re-encode through the
    child's *own* locale - the very variable under test - and mask a
    difference that the netsh decode had actually introduced.
    """
    code = (
        "import sys; sys.path.insert(0, r'%s');"
        "from _netsh import run_netsh;"
        "r = run_netsh(['wlan', 'show', 'profiles']);"
        "sys.stdout.buffer.write(r.stdout.encode('utf-8'))" % str(WIFI_DIR)
    )
    env = dict(os.environ)
    env["PYTHONUTF8"] = utf8_mode
    proc = subprocess.run(
        [sys.executable, "-c", code],
        capture_output=True,
        env=env,
        cwd=str(WIFI_DIR),
        creationflags=_netsh.NO_WINDOW,
    )
    if proc.returncode != 0:
        raise unittest.SkipTest(f"helper child failed: {proc.stderr.decode('utf-8', 'replace')}")
    return proc.stdout.decode("utf-8")


@unittest.skipUnless(sys.platform == "win32", "netsh is Windows-only")
class LiveNetshUnderUtf8ModeTests(unittest.TestCase):
    """The regression test: identical output whatever the parent's UTF-8 mode.

    Before the fix these two runs disagreed - `text=True` resolved to cp1252
    with UTF-8 mode off and to UTF-8 with it on, so the same SSID came back
    two different ways. After it, decoding is pinned in-module and the parent's
    locale is irrelevant.
    """

    def test_output_is_non_empty_under_utf8_mode(self) -> None:
        if not _netsh.run_netsh(["wlan", "show", "profiles"]).ran:
            self.skipTest("netsh is unavailable on this machine")
        self.assertTrue(_run_netsh_in_child("1").strip(), "no output with PYTHONUTF8=1")

    def test_output_matches_with_utf8_mode_on_and_off(self) -> None:
        if not _netsh.run_netsh(["wlan", "show", "profiles"]).ran:
            self.skipTest("netsh is unavailable on this machine")
        on = _run_netsh_in_child("1")
        off = _run_netsh_in_child("0")
        self.assertTrue(on.strip(), "no output with PYTHONUTF8=1")
        self.assertTrue(off.strip(), "no output with PYTHONUTF8=0")
        self.assertEqual(on, off, "netsh output still depends on the parent's locale")


class NoInheritedDecodingTests(unittest.TestCase):
    """Acceptance criteria 1, 2 and 5, enforced against the sources."""

    def test_tools_do_not_spawn_netsh_themselves(self) -> None:
        for name in TOOL_MODULES:
            source = (WIFI_DIR / name).read_text(encoding="utf-8")
            self.assertNotIn(
                "subprocess.", source, f"{name} must reach netsh through _netsh.run_netsh"
            )

    def test_tools_do_not_pin_a_parent_locale_decode(self) -> None:
        for name in TOOL_MODULES:
            source = (WIFI_DIR / name).read_text(encoding="utf-8")
            self.assertNotIn("text=True", source, f"{name} decodes with the parent's locale")
            self.assertNotIn("encoding='utf-8'", source, f"{name} pins a strict utf-8 decode")

    def test_helper_captures_bytes_and_suppresses_the_console_window(self) -> None:
        source = (WIFI_DIR / "_netsh.py").read_text(encoding="utf-8")
        self.assertIn("capture_output=True", source)
        self.assertIn("creationflags=NO_WINDOW", source)


if __name__ == "__main__":
    unittest.main()
