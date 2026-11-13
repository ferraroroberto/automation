#!/usr/bin/env python3
r"""Decoding native Windows console-tool output, shared across the repo.

Why this module exists
----------------------
`subprocess.run(..., text=True)` decodes the child's output with the
*parent's* ambient locale, never the child's console code page. Native
Windows console tools (`netsh`, `pnputil`, `nvidia-smi`, `powershell`,
`schtasks`, `sc`, `tasklist`, `reg`, `ipconfig`, ...) write their own
console page, so the two disagree the moment the parent runs in Python's
UTF-8 mode - `PYTHONUTF8=1`, or a parent that enables it: a tray app, a
wrapper .bat, an app-launcher job, an agent session capturing output.

Measured on this machine (Windows 11, OEM code page 850, ANSI 1252) against
a child emitting the cp850 bytes ``Configuraci\xa2n`` - both halves of the
mismatch bite, and neither is what a reader expects:

* `PYTHONUTF8` unset - `text=True` decodes with cp1252 and silently hands
  back the **wrong character** (``Configuraci¢n`` for ``Configuración``).
  Exit code 0, plausible-looking string, corrupt data.
* `PYTHONUTF8=1` - the `UnicodeDecodeError` is raised inside subprocess's
  *reader thread*, which dumps a traceback and dies. `CompletedProcess`
  then carries `returncode == 0` and **`stdout is None`** - not `""`, and
  not a truncated read: the whole stream is gone. The caller's
  `result.stdout.strip()` therefore raises **`AttributeError`**, which
  `except (subprocess.SubprocessError, OSError)` does not catch, so the
  tool dies with an unhandled traceback well away from the real cause.

A silently blank or missing result from a scheduling/status query (e.g. a
`next_run` field) is indistinguishable from a genuinely quiet system until
someone checks the raw bytes - the rule below is the fix for that failure
mode wherever a native console tool's output feeds a decision.

So every caller here captures raw bytes - never `text=`, never `encoding=`
- and decodes explicitly through `decode_console_bytes`.

Why a candidate list rather than one pinned codec
-------------------------------------------------
The usual fleet remedy is to pin `encoding="oem", errors="replace"`, on the
grounds that these tools emit the OEM code page. Measured for `netsh`,
that is only half true: profile names come back as **UTF-8**
regardless of the console page - the SSID `Roberto’s iPhone` arrives as the
bytes `e2 80 99` - so a pinned `oem` decode turns it into
`RobertoÔÇÖs iPhone`. netsh's own localized labels *do* follow the console
page, so one stream can legitimately mix both encodings.

Hence the ordered list below: UTF-8 first (correct for the measured case),
then the OEM console page (correct for localized labels on a stream that is
not valid UTF-8), then ANSI, and finally a lossy UTF-8 pass so that one odd
byte costs a character rather than the whole result. Decoding therefore
never raises and never silently yields "" for non-empty output.

Importing this module
---------------------
There is no repo-wide package, so reach it the way the other cross-folder
helpers here are reached - put `tools/system/` on `sys.path` and import::

    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "system"))
    from _lib.console_output import NO_WINDOW, decode_console_bytes
"""

from __future__ import annotations

import logging
import subprocess
import sys
from pathlib import Path
from typing import NamedTuple, Optional, Sequence, Tuple

logger = logging.getLogger(__name__)

# Ordered decode candidates - see the module docstring for why each is here.
# "oem"/"mbcs" are Windows-only codec aliases, so POSIX gets UTF-8 alone.
CONSOLE_ENCODINGS: Tuple[str, ...] = (
    ("utf-8", "oem", "mbcs") if sys.platform == "win32" else ("utf-8",)
)

# Suppress the console window each spawn would otherwise flash on a parent
# that has no console of its own (pythonw, a tray app, a scheduled task).
# Re-exported from the repo-root helper so the two never drift.
sys.path.append(str(Path(__file__).resolve().parents[3]))
from no_window import NO_WINDOW as NO_WINDOW  # noqa: E402


def decode_console_bytes(raw: bytes) -> Tuple[str, str]:
    """Decode native-console-tool output.

    Returns the decoded text and the codec that produced it. Never raises:
    the final candidate decodes with ``errors="replace"``, so unexpected
    bytes cost a character rather than the whole output.
    """
    for encoding in CONSOLE_ENCODINGS:
        try:
            return raw.decode(encoding), encoding
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode("utf-8", errors="replace"), "utf-8/replace"


class ConsoleRun(NamedTuple):
    """Outcome of one native-console-tool invocation.

    `ran` is the state that must never be folded into a negative answer: a
    query that could not run, or whose output could not be read, is
    *unknown* - not "the device is gone". `error` says why it did not run.
    """

    ran: bool
    returncode: int
    stdout: str
    stderr: str
    encoding: Optional[str]
    error: Optional[str] = None

    @property
    def succeeded(self) -> bool:
        """True when the tool ran and reported success."""
        return self.ran and self.returncode == 0

    @property
    def has_output(self) -> bool:
        """True when stdout carries something to parse."""
        return bool(self.stdout.strip())


def run_console(
    command: Sequence[str], *, timeout: int, label: str, strict: bool = False
) -> ConsoleRun:
    """Run a native Windows console tool and decode its output explicitly.

    Captures raw bytes - never `text=`, never `encoding=` - because those
    decode with the *parent's* locale rather than the child's console code
    page (see the module docstring). Never raises: every failure path logs
    and comes back as a `ConsoleRun`, so an unreadable result leaves a
    breadcrumb instead of a silent negative.

    `strict` is for a tool that always prints on success (netsh): a non-zero
    exit then logs a warning and blank output on exit 0 logs an error. Left
    off, a non-zero exit is debug-level (the caller treats it as an answer)
    and only unreadable output - raw bytes that decode to nothing - is an
    error, since a PowerShell snippet may legitimately print nothing.
    """
    try:
        proc = subprocess.run(
            command,
            capture_output=True,
            creationflags=NO_WINDOW,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        message = f"timed out after {timeout}s"
        logger.error("❌ %s %s", label, message)
        return ConsoleRun(False, -1, "", "", None, message)
    except (subprocess.SubprocessError, OSError) as exc:
        logger.error("❌ Could not run %s: %s", label, exc)
        return ConsoleRun(False, -1, "", "", None, str(exc))

    stdout, encoding = decode_console_bytes(proc.stdout or b"")
    stderr, _ = decode_console_bytes(proc.stderr or b"")

    if proc.returncode != 0:
        log_nonzero = logger.warning if strict else logger.debug
        log_nonzero("ℹ️ %s exited %d: %s", label, proc.returncode, (stderr or stdout).strip())
    elif not stdout.strip() and (strict or proc.stdout):
        # The answer was not established - never treat it as a real "none".
        logger.error(
            "❌ %s exited 0 but produced no readable output "
            "(%d raw bytes, decoded as %s) - result is unknown, not empty.",
            label,
            len(proc.stdout or b""),
            encoding,
        )

    return ConsoleRun(True, proc.returncode, stdout, stderr, encoding)
