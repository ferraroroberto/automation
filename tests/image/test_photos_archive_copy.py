"""copy_files must never mark an uncopied source as copied (it is deleted afterwards)."""

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "image"))

import photos_archive  # noqa: E402


def _frame(source_path: str, unified_name: str = "20240115-143045.jpg") -> pd.DataFrame:
    return pd.DataFrame([{
        "file_path": source_path,
        "unified_name": unified_name,
        "file_extension": ".jpg",
        "duplicate_count": 1,
        "discard": 0,
        "exclude": 0,
        "copy_success": 0,
        "destination_path": "",
    }])


class CopyFilesCollisionTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.source = Path(self._tmp.name, "inbox")
        self.archive = Path(self._tmp.name, "archive")
        self.source.mkdir()
        self.archive.mkdir()
        patcher = mock.patch.object(photos_archive, "get_config", return_value={"destination_folder": str(self.archive)})
        patcher.start()
        self.addCleanup(patcher.stop)

    def _write(self, folder: Path, name: str, data: bytes) -> str:
        path = folder / name
        path.write_bytes(data)
        return str(path)

    def test_a_different_photo_with_the_same_name_is_copied_under_the_next_suffix(self):
        self._write(self.archive, "20240115-143045.jpg", b"first camera")
        source = self._write(self.source, "IMG_new.jpg", b"second camera")
        df = _frame(source)

        photos_archive.copy_files(df)

        self.assertEqual(df.at[0, "copy_success"], 1)
        self.assertEqual(Path(df.at[0, "destination_path"]).name, "20240115-143045-001.jpg")
        self.assertEqual(Path(df.at[0, "destination_path"]).read_bytes(), b"second camera")
        self.assertEqual((self.archive / "20240115-143045.jpg").read_bytes(), b"first camera")

    def test_an_identical_file_already_in_the_archive_counts_as_copied_without_a_second_copy(self):
        self._write(self.archive, "20240115-143045.jpg", b"same bytes")
        source = self._write(self.source, "IMG_new.jpg", b"same bytes")
        df = _frame(source)

        photos_archive.copy_files(df)

        self.assertEqual(df.at[0, "copy_success"], 1)
        self.assertEqual(sorted(os.listdir(self.archive)), ["20240115-143045.jpg"])

    def test_a_free_name_is_used_as_is(self):
        source = self._write(self.source, "IMG_new.jpg", b"only")
        df = _frame(source)

        photos_archive.copy_files(df)

        self.assertEqual(Path(df.at[0, "destination_path"]).name, "20240115-143045.jpg")


if __name__ == "__main__":
    unittest.main()
