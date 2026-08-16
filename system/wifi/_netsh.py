#!/usr/bin/env python3
r"""Shared `netsh` runner for the Wi-Fi tools in this folder.

Why this module exists
----------------------
`subprocess.run(..., text=True)` decodes the child's output with the
*parent's* ambient locale. Under Python's UTF-8 mode (`PYTHONUTF8=1`, or a
parent that enables it — a tray app, a wrapper .bat, an app-launcher job)
that is UTF-8, while native Windows console tools write their own console
code page. The mismatch is not reliably loud: the output comes back
mangled or empty instead of raising, and a caller that inspects it then
reports "not connected" / "no profiles" / "export failed" for reasons that
have nothing to do with the network. `app-launcher#743` is the fleet
incident behind the rule; `encoding='utf-8'` is the same bug stated
explicitly rather than inherited.

So every call here captures raw bytes — never `text=`, never `encoding=` —
and decodes explicitly.

The decoder itself, and the measured rationale for its ordered candidate
list, moved to `system/_lib/console_output.py` in `#108` so the second
instance of this bug (`video/gpu_recovery.py`) could reuse it rather than
re-derive it. The names below are re-exported unchanged: this module's
callers and its test suite still reach `decode_console_bytes`, `NO_WINDOW`
and `NETSH_ENCODINGS` through `_netsh`.
"""

from __future__ import annotations

import logging
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import NamedTuple, Optional, Sequence, Tuple

# _lib/ lives two levels up (system/_lib/), alongside this package's parent.
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from _lib.console_output import (  # noqa: E402
    CONSOLE_ENCODINGS,
    NO_WINDOW,
    decode_console_bytes,
)

logger = logging.getLogger(__name__)

# Re-exported under this module's historical name (see the docstring).
NETSH_ENCODINGS: Tuple[str, ...] = CONSOLE_ENCODINGS

__all__ = [
    "CurrentSsid",
    "NETSH_ENCODINGS",
    "NO_WINDOW",
    "NetshResult",
    "current_ssid",
    "decode_console_bytes",
    "parse_ssid",
    "run_netsh",
]


@dataclass(frozen=True)
class NetshResult:
    """Outcome of one `netsh` invocation.

    `ran` is the state that must never be folded into a negative answer: a
    query that could not run at all is *unknown*, not "no networks".
    """

    args: Tuple[str, ...]
    returncode: int
    stdout: str
    stderr: str
    encoding: Optional[str]
    error: Optional[str] = None

    @property
    def ran(self) -> bool:
        """True when netsh actually executed (whatever its exit code)."""
        return self.error is None

    @property
    def succeeded(self) -> bool:
        """True when netsh ran and reported success."""
        return self.ran and self.returncode == 0

    @property
    def has_output(self) -> bool:
        """True when stdout carries something to parse."""
        return bool(self.stdout.strip())

    @property
    def combined(self) -> str:
        """stdout + stderr, for callers that report either to the user."""
        return self.stdout + self.stderr


def run_netsh(args: Sequence[str]) -> NetshResult:
    """Run `netsh <args>` and return a decoded, state-carrying result.

    Never raises for a missing or failing netsh — the failure is logged and
    exposed via `NetshResult.ran` / `.succeeded` so callers can tell "could
    not establish this" apart from "established, and the answer is nothing".
    """
    command = ["netsh", *args]
    printable = " ".join(command)
    try:
        proc = subprocess.run(
            command,
            capture_output=True,
            creationflags=NO_WINDOW,
        )
    except (FileNotFoundError, OSError) as exc:
        logger.error("❌ Could not run '%s': %s", printable, exc)
        return NetshResult(
            args=tuple(command),
            returncode=-1,
            stdout="",
            stderr="",
            encoding=None,
            error=str(exc),
        )

    stdout, encoding = decode_console_bytes(proc.stdout or b"")
    stderr, _ = decode_console_bytes(proc.stderr or b"")

    if proc.returncode != 0:
        logger.warning(
            "⚠️ '%s' exited %d: %s", printable, proc.returncode, (stderr or stdout).strip()
        )
    elif not stdout.strip():
        # netsh always prints something on success, so an empty stdout means
        # the answer was not established — never treat it as a real "none".
        logger.error(
            "❌ '%s' exited 0 but produced no readable output "
            "(%d raw bytes, decoded as %s) - result is unknown, not empty.",
            printable,
            len(proc.stdout or b""),
            encoding,
        )
    else:
        logger.debug("ℹ️ '%s' decoded %d chars as %s", printable, len(stdout), encoding)

    return NetshResult(
        args=tuple(command),
        returncode=proc.returncode,
        stdout=stdout,
        stderr=stderr,
        encoding=encoding,
    )


def parse_ssid(output: str) -> Optional[str]:
    """Pull the connected SSID out of `netsh wlan show interfaces` output.

    Returns None when the output holds no SSID line — i.e. not connected.
    """
    for line in output.splitlines():
        # The SSID line contains "SSID" but not "BSSID"; the value follows the colon.
        if "SSID" in line and "BSSID" not in line:
            _, _, value = line.partition(":")
            ssid = value.strip()
            if ssid:
                return ssid
    return None


class CurrentSsid(NamedTuple):
    """Tri-state answer to "which network are we on?".

    `known` is False when the netsh query itself failed, in which case
    `ssid` is meaningless — that case must not be shown as "not connected".
    """

    ssid: Optional[str]
    known: bool


def current_ssid() -> CurrentSsid:
    """Return the SSID currently connected to, distinguishing unknown from none."""
    result = run_netsh(["wlan", "show", "interfaces"])
    if not result.ran or not result.has_output:
        return CurrentSsid(ssid=None, known=False)
    return CurrentSsid(ssid=parse_ssid(result.stdout), known=True)
