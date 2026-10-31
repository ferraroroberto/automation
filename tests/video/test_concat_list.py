"""The concat-demuxer list file must survive quotes and non-cp1252 names."""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "video"))

import video_concatenator  # noqa: E402


class ConcatListTextTests(unittest.TestCase):
    def test_plain_path(self):
        self.assertEqual(video_concatenator._concat_list_text(["C:/v/a.mp4"]), "file 'C:/v/a.mp4'\n")

    def test_single_quote_is_escaped_for_the_demuxer(self):
        text = video_concatenator._concat_list_text(["C:/v/Mom's party.mp4"])
        self.assertEqual(text, "file 'C:/v/Mom'\\''s party.mp4'\n")

    def test_one_line_per_clip_in_order(self):
        text = video_concatenator._concat_list_text(["/a.mp4", "/b.mp4"])
        self.assertEqual(text.splitlines(), ["file '/a.mp4'", "file '/b.mp4'"])


if __name__ == "__main__":
    unittest.main()
