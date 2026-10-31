"""A second launch hands its .md path to the running instance (no window, no Windows APIs)."""

import tempfile
import unittest
from pathlib import Path

from system import markdown_preview as mp


class HandoffTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.handoff = Path(self._tmp.name) / "open.txt"

    def test_path_written_by_the_second_launch_is_taken_once(self):
        target = Path(self._tmp.name) / "notes.md"
        mp.write_handoff(target, self.handoff)
        self.assertEqual(mp.take_handoff(self.handoff), target)
        self.assertIsNone(mp.take_handoff(self.handoff))

    def test_no_handoff_file_means_nothing_to_load(self):
        self.assertIsNone(mp.take_handoff(self.handoff))

    def test_blank_handoff_is_ignored(self):
        self.handoff.write_text("  ", encoding="utf-8")
        self.assertIsNone(mp.take_handoff(self.handoff))

    def test_second_launch_hands_over_its_file_before_signalling(self):
        # _acquire_lock writes the hand-off only when it was given a file.
        source = Path(mp.__file__).read_text(encoding="utf-8")
        self.assertIn("write_handoff(self.current_file.resolve())", source)
        self.assertLess(source.index("write_handoff(self.current_file"), source.index("k.SetEvent(h)"))


if __name__ == "__main__":
    unittest.main()
