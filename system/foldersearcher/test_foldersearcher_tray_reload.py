#!/usr/bin/env python3
"""Tests for the tray app picking up an index rebuilt behind its back.

Exercises ``FolderSearcher.reload_structure_if_changed`` without starting the
tray icon or a Tk root: the instance is built with ``__new__`` and only the
attributes the load path touches are set.

Run from the repo root:
    & .\\.venv\\Scripts\\python.exe -m unittest discover -s system/foldersearcher -p "test_*.py"
"""

import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import foldersearcher  # noqa: E402
import foldersearcher_core as core  # noqa: E402


class _StatusStub:
    def set(self, value):
        self.value = value


def write_index(path, root, *folders):
    index = core.FolderIndex([root])
    index.entries = {root: [f"{root}/{name}" for name in folders]}
    index.save(path)


class TestTrayReloadsNewerIndex(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.structure_file = os.path.join(self._tmp.name, "folder_structure.txt")

        self.app = foldersearcher.FolderSearcher.__new__(foldersearcher.FolderSearcher)
        self.app.config = core.FolderSearcherConfig(root_paths=["X:/docs"], structure_file=self.structure_file)
        self.app.index = core.FolderIndex()
        self.app._index_mtime = None
        self.app.status_var = _StatusStub()

    def tearDown(self):
        self._tmp.cleanup()

    def bump_mtime(self, seconds):
        stat = os.stat(self.structure_file)
        os.utime(self.structure_file, (stat.st_atime, stat.st_mtime + seconds))

    def test_rewritten_index_is_reloaded_once(self):
        write_index(self.structure_file, "X:/docs", "old")
        self.app.load_structure()
        self.assertIn("X:/docs/old", self.app.index.all_paths())
        self.assertFalse(self.app.reload_structure_if_changed())

        write_index(self.structure_file, "X:/docs", "created-today")
        self.bump_mtime(10)

        self.assertTrue(self.app.reload_structure_if_changed())
        self.assertIn("X:/docs/created-today", self.app.index.all_paths())
        self.assertNotIn("X:/docs/old", self.app.index.all_paths())
        self.assertFalse(self.app.reload_structure_if_changed())

    def test_index_appearing_after_startup_is_loaded(self):
        self.app.load_structure()
        self.assertEqual(len(self.app.index), 0)

        write_index(self.structure_file, "X:/docs", "first-scan")

        self.assertTrue(self.app.reload_structure_if_changed())
        self.assertIn("X:/docs/first-scan", self.app.index.all_paths())

    def test_no_index_file_is_not_a_reload(self):
        self.assertFalse(self.app.reload_structure_if_changed())


if __name__ == "__main__":
    unittest.main()
