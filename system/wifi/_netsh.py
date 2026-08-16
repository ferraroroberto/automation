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

Why a candidate list rather than one pinned codec
-------------------------------------------------
The usual fleet remedy is to pin `encoding="oem", errors="replace"`, on the
grounds that these tools emit the OEM code page. Measured on this machine
(Windows 11, `GetConsoleOutputCP()` == `GetOEMCP()` == 850), that is only
half true for `netsh`: profile names come back as **UTF-8** regardless of
the console page — the SSID `Roberto’s iPhone` arrives as the bytes
`e2 80 99` — so a pinned `oem` decode turns it into `RobertoÔÇÖs iPhone`.
netsh's own localized labels *do* follow the console page, so one stream
can legitimately mix both encodings.

Hence the ordered list below: UTF-8 first (correct for the measured case),
then the OEM console page (correct for localized labels on a stream that is
not valid UTF-8), then ANSI, and finally a lossy UTF-8 pass so that one odd
byte costs a character rather than the whole result. Decoding therefore
never raises and never silently yields "" for non-empty output.
"""

from __future__ import annotations

import logging
import subprocess
import sys
from dataclasses import dataclass
from typing import NamedTuple, Optional, Sequence, Tuple

logger = logging.getLogger(__name__)

# Ordered decode candidates — see the module docstring for why each is here.
# "oem"/"mbcs" are Windows-only codec aliases, so POSIX gets UTF-8 alone.
NETSH_ENCODINGS: Tuple[str, ...] = (
    ("utf-8", "oem", "mbcs") if sys.platform == "win32" else ("utf-8",)
)

# Suppress the console window each netsh spawn would otherwise flash.
NO_WINDOW = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0


def decode_console_bytes(raw: bytes) -> Tuple[str, str]:
    """Decode native-console-tool output.

    Returns the decoded text and the codec that produced it. Never raises:
    the final candidate decodes with ``errors="replace"``, so unexpected
    bytes cost a character rather than the whole output.
    """
    for encoding in NETSH_ENCODINGS:
        try:
            return raw.decode(encoding), encoding
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode("utf-8", errors="replace"), "utf-8/replace"


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
