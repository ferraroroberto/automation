#!/usr/bin/env python3
r"""Guard: a recording failure must reach the user as a "Recording Error" dialog.

`_record` runs on a worker thread and defers the dialog to the Tk loop with
`master.after(0, lambda: ...)`. Python unbinds `except ... as e` when the block
ends, so a lambda that reads `e` raises NameError once the Tk loop runs it and
the failure message is never shown.

Run from the repo root:

    & .\.venv\Scripts\python.exe -m unittest discover -s tests/video -p "test_*.py"
"""

from __future__ import annotations

import sys
import threading
import unittest
from pathlib import Path
from unittest import mock

REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO_ROOT / "video"))

import screen_recorder  # noqa: E402


class _DeferredMaster:
    """Stands in for Tk: `after` queues callbacks to run after `_record` returns."""

    def __init__(self) -> None:
        self.queued: list = []

    def after(self, _delay_ms: int, callback) -> None:
        self.queued.append(callback)


class RecordErrorDialogTests(unittest.TestCase):
    def test_capture_failure_shows_recording_error_with_exception_text(self) -> None:
        recorder = object.__new__(screen_recorder.ScreenRecorderGUI)
        recorder.master = _DeferredMaster()
        recorder._recording_event = threading.Event()
        recorder._recording_event.set()
        recorder._fps = 5
        recorder.stop_recording = mock.Mock()

        with mock.patch.object(screen_recorder.mss, "mss", side_effect=RuntimeError("capture boom")), \
                mock.patch.object(screen_recorder.messagebox, "showerror") as showerror:
            recorder._record()
            for callback in recorder.master.queued:
                callback()

        showerror.assert_called_once_with("Recording Error", "capture boom")
        recorder.stop_recording.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
