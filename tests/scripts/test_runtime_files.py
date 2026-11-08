"""`system/_lib/runtime_files` writes caller-named paths and refuses any path inside the checkout."""

from __future__ import annotations

import logging
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.append(str(REPO_ROOT))

from system._lib import runtime_files  # noqa: E402


class OutsideRepoTests(unittest.TestCase):
    def test_refuses_a_path_inside_the_checkout(self) -> None:
        with self.assertRaises(ValueError):
            runtime_files.outside_repo(REPO_ROOT / "audio" / "out.txt")

    def test_refuses_the_checkout_itself(self) -> None:
        with self.assertRaises(ValueError):
            runtime_files.outside_repo(REPO_ROOT)

    def test_accepts_a_temp_path(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(runtime_files.outside_repo(tmp), Path(tmp).resolve())


class WriterTests(unittest.TestCase):
    def test_writers_create_and_fill_files_outside_the_checkout(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = runtime_files.make_dirs(Path(tmp) / "a" / "b")
            runtime_files.save_text(folder / "t.txt", "hi")
            runtime_files.save_bytes(folder / "b.bin", b"\x00\x01")
            runtime_files.copy_file(folder / "t.txt", folder / "copy.txt")
            self.assertEqual((folder / "t.txt").read_text(encoding="utf-8"), "hi")
            self.assertEqual((folder / "b.bin").read_bytes(), b"\x00\x01")
            self.assertEqual((folder / "copy.txt").read_text(encoding="utf-8"), "hi")

    def test_writers_refuse_the_checkout(self) -> None:
        inside = REPO_ROOT / "runtime_files_test_probe"
        for call in (
            lambda: runtime_files.make_dirs(inside),
            lambda: runtime_files.save_text(inside / "x.txt", "x"),
            lambda: runtime_files.open_for_write(inside / "x.txt"),
            lambda: runtime_files.copy_file(__file__, inside / "x.txt"),
            lambda: runtime_files.file_handler(inside / "x.log"),
        ):
            with self.assertRaises(ValueError):
                call()
        self.assertFalse(inside.exists())

    def test_file_handler_creates_its_folder(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            handler = runtime_files.file_handler(Path(tmp) / "logs" / "x.log")
            try:
                self.assertIsInstance(handler, logging.FileHandler)
                self.assertTrue((Path(tmp) / "logs").is_dir())
            finally:
                handler.close()


class RuntimeDataDirTests(unittest.TestCase):
    def test_env_override_wins_and_must_be_outside_the_checkout(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            with mock.patch.dict(os.environ, {runtime_files.RUNTIME_DIR_ENV: tmp}):
                self.assertEqual(runtime_files.runtime_data_dir(), Path(tmp).resolve())
        with mock.patch.dict(os.environ, {runtime_files.RUNTIME_DIR_ENV: str(REPO_ROOT / "state")}):
            with self.assertRaises(ValueError):
                runtime_files.runtime_data_dir()

    def test_default_is_outside_the_checkout(self) -> None:
        with mock.patch.dict(os.environ, {}, clear=False):
            os.environ.pop(runtime_files.RUNTIME_DIR_ENV, None)
            self.assertNotIn(REPO_ROOT, runtime_files.runtime_data_dir().parents)


if __name__ == "__main__":
    unittest.main()
