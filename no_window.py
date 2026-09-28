"""Shared ``creationflags`` value that keeps subprocess spawns from flashing a console.

Every ``subprocess`` spawn of an external executable passes
``creationflags=NO_WINDOW`` so a parent with no console of its own (pythonw, a
tray app, a Tk GUI launched from a ``.bat``, a scheduled task) does not pop a
console window per child. It is ``0`` off Windows, so call sites need no
platform check of their own.

Scripts here are run as ``python <folder>/<script>.py``, so ``sys.path[0]`` is
the script's folder, not the repo root. A consumer therefore appends the repo
root before importing::

    sys.path.append(str(Path(__file__).resolve().parents[1]))
    from no_window import NO_WINDOW  # noqa: E402

(``parents[1]`` for a script one folder below the root; adjust per depth.)
Never re-derive the ternary at a call site - import this constant.
"""

from __future__ import annotations

import subprocess
import sys

NO_WINDOW: int = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0

__all__ = ["NO_WINDOW"]
