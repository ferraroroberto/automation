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

`app-launcher#743` is the fleet incident behind the rule (blank `next_run`
on all 20 jobs for weeks, because a dead query was indistinguishable from a
quiet system); `#106` fixed the `netsh` instance in this repo and `#108`
the `video/gpu_recovery.py` one.

So every caller here captures raw bytes - never `text=`, never `encoding=`
- and decodes explicitly through `decode_console_bytes`.

Why a candidate list rather than one pinned codec
-------------------------------------------------
The usual fleet remedy is to pin `encoding="oem", errors="replace"`, on the
grounds that these tools emit the OEM code page. Measured for `netsh`
(`#106`), that is only half true: profile names come back as **UTF-8**
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
helpers here are reached - put `system/` on `sys.path` and import::

    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "system"))
    from _lib.console_output import NO_WINDOW, decode_console_bytes
"""

from __future__ import annotations

import subprocess
import sys
from typing import Tuple

# Ordered decode candidates - see the module docstring for why each is here.
# "oem"/"mbcs" are Windows-only codec aliases, so POSIX gets UTF-8 alone.
CONSOLE_ENCODINGS: Tuple[str, ...] = (
    ("utf-8", "oem", "mbcs") if sys.platform == "win32" else ("utf-8",)
)

# Suppress the console window each spawn would otherwise flash on a parent
# that has no console of its own (pythonw, a tray app, a scheduled task).
NO_WINDOW = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0


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
