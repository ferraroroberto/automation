"""Thin launcher so the tool runs without `python -m …` from this folder.

Usage:
    python launcher.py tray
    python launcher.py record
    python launcher.py gui
    python launcher.py server start

The launcher just patches `sys.path` so the `audio.transcribe_voice`
package is importable when the script is run from inside its own folder.
"""

from __future__ import annotations

import sys
from pathlib import Path

# Add repo root (three levels up) to sys.path so `audio.transcribe_voice...` resolves.
REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from audio.transcribe_voice.cli.main import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())
