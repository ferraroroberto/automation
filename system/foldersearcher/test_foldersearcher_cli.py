#!/usr/bin/env python3
"""Tests for the headless scan CLI (foldersearcher_cli).

Run from the repo root:
    & .\\.venv\\Scripts\\python.exe -m unittest discover -s system/foldersearcher -p "test_*.py"
"""

import json
import os
import sys
import tempfile
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import foldersearcher_cli as cli  # noqa: E402
import foldersearcher_core as core  # noqa: E402

OLD_INDEX = "Root: X:/old\nX:/old\n  X:/old/kept\n\n"


class CliScanTestBase(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = self._tmp.name
        self.config_file = os.path.join(self.tmp, "foldersearcher.json")
        self.structure_file = os.path.join(self.tmp, "folder_structure.txt")
        with open(self.structure_file, "w", encoding="utf-8") as handle:
            handle.write(OLD_INDEX)

    def tearDown(self):
        self._tmp.cleanup()

    def make_root(self, name, *children):
        root = os.path.join(self.tmp, name)
        for child in children:
            os.makedirs(os.path.join(root, child))
        os.makedirs(root, exist_ok=True)
        return root

    def write_config(self, data):
        with open(self.config_file, "w", encoding="utf-8") as handle:
            json.dump(data, handle)
        with open(self.config_file, "rb") as handle:
            return handle.read()

    def read_index_bytes(self):
        with open(self.structure_file, "rb") as handle:
            return handle.read()

    def leftover_temp_files(self):
        return [name for name in os.listdir(self.tmp) if name.endswith(".tmp")]


class TestScanSuccess(CliScanTestBase):
    def test_rebuilds_index_in_sectioned_format(self):
        alpha = self.make_root("alpha", "clientes/acme", "archive")
        beta = self.make_root("beta", "work")
        self.write_config({"root_paths": [alpha, beta]})

        code = cli.run_scan(self.config_file, self.structure_file)

        self.assertEqual(code, cli.EXIT_OK)
        index = core.FolderIndex()
        index.load(self.structure_file)
        self.assertEqual(index.roots, core.normalize_roots([alpha, beta]))
        self.assertIn(core.normalize_root(os.path.join(alpha, "clientes", "acme")), index.all_paths())
        self.assertIn(core.normalize_root(os.path.join(beta, "work")), index.all_paths())
        self.assertNotIn("X:/old/kept", index.all_paths())
        self.assertEqual(self.leftover_temp_files(), [])

    def test_legacy_root_folder_config_is_scanned_but_not_rewritten(self):
        alpha = self.make_root("alpha", "sub")
        before = self.write_config({"root_folder": alpha, "structure_file": "folder_structure.txt"})

        code = cli.run_scan(self.config_file, self.structure_file)

        self.assertEqual(code, cli.EXIT_OK)
        with open(self.config_file, "rb") as handle:
            self.assertEqual(handle.read(), before)


class TestScanRefusesToShrinkIndex(CliScanTestBase):
    def test_missing_root_exits_non_zero_and_leaves_index_byte_identical(self):
        alpha = self.make_root("alpha", "sub")
        self.write_config({"root_paths": [alpha, os.path.join(self.tmp, "unplugged")]})
        before = self.read_index_bytes()

        code = cli.run_scan(self.config_file, self.structure_file)

        self.assertEqual(code, cli.EXIT_MISSING_ROOT)
        self.assertEqual(self.read_index_bytes(), before)

    def test_no_roots_exits_non_zero_and_leaves_index_byte_identical(self):
        self.write_config({"root_paths": []})
        before = self.read_index_bytes()

        code = cli.run_scan(self.config_file, self.structure_file)

        self.assertEqual(code, cli.EXIT_NO_ROOTS)
        self.assertEqual(self.read_index_bytes(), before)

    def test_missing_config_file_counts_as_no_roots(self):
        before = self.read_index_bytes()

        code = cli.run_scan(self.config_file, self.structure_file)

        self.assertEqual(code, cli.EXIT_NO_ROOTS)
        self.assertEqual(self.read_index_bytes(), before)


class TestAtomicWrite(CliScanTestBase):
    def test_failure_mid_write_leaves_previous_index_and_no_temp_file(self):
        alpha = self.make_root("alpha", "sub")
        self.write_config({"root_paths": [alpha]})
        before = self.read_index_bytes()

        def partial_save(index, path):
            with open(path, "w", encoding="utf-8") as handle:
                handle.write("Root: half-written")
            raise OSError("disk full")

        with mock.patch.object(core.FolderIndex, "save", partial_save):
            code = cli.run_scan(self.config_file, self.structure_file)

        self.assertEqual(code, cli.EXIT_WRITE_FAILED)
        self.assertEqual(self.read_index_bytes(), before)
        self.assertEqual(self.leftover_temp_files(), [])


class TestMainEntryPoint(unittest.TestCase):
    def test_scan_subcommand_uses_script_dir_files(self):
        with mock.patch.object(cli, "run_scan", return_value=cli.EXIT_OK) as run_scan:
            self.assertEqual(cli.main(["scan"]), cli.EXIT_OK)
        run_scan.assert_called_once_with(cli.DEFAULT_CONFIG_FILE, cli.DEFAULT_STRUCTURE_FILE)


if __name__ == "__main__":
    unittest.main()
