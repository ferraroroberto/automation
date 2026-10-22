"""Regression tests for the photo-archive destination-inside-source checks."""

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "image"))

import photos_archive_metadata  # noqa: E402


def _forward_slashes(path: str) -> str:
    # The form photos_archive.json and photos_archive.md use.
    return path.replace(os.sep, "/")


class IsInsideFolderTests(unittest.TestCase):
    def test_forward_slash_config_values_are_recognised(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = _forward_slashes(tmp)
            dest = _forward_slashes(os.path.join(tmp, "archive"))
            self.assertTrue(photos_archive_metadata.is_inside_folder(dest, source))

    def test_folder_counts_as_inside_itself(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertTrue(photos_archive_metadata.is_inside_folder(tmp, _forward_slashes(tmp)))

    def test_sibling_with_shared_prefix_is_not_inside(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = os.path.join(tmp, "photos")
            sibling = os.path.join(tmp, "photos_archive")
            self.assertFalse(photos_archive_metadata.is_inside_folder(sibling, source))

    def test_unrelated_folders_are_not_inside(self):
        with tempfile.TemporaryDirectory() as a, tempfile.TemporaryDirectory() as b:
            self.assertFalse(photos_archive_metadata.is_inside_folder(a, b))


class ProcessFilesSkipDestTests(unittest.TestCase):
    def test_destination_inside_source_is_skipped_with_forward_slash_dest(self):
        with tempfile.TemporaryDirectory() as tmp:
            dest = os.path.join(tmp, "archive")
            os.makedirs(dest)
            for folder in (tmp, dest):
                with open(os.path.join(folder, "pic.jpg"), "w") as handle:
                    handle.write("x")

            scanned = []

            def fake_info(path):
                scanned.append(os.path.normpath(path))
                return {"path": path}

            with mock.patch.object(photos_archive_metadata, "get_file_info", side_effect=fake_info):
                photos_archive_metadata.process_files(tmp, _forward_slashes(dest), True)

            self.assertEqual(scanned, [os.path.normpath(os.path.join(tmp, "pic.jpg"))])


if __name__ == "__main__":
    unittest.main()
