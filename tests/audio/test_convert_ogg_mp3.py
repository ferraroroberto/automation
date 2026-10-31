"""convert_ogg_mp3 runs the bundled ffmpeg directly; pydub's ffprobe step is gone."""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "audio"))

import convert_ogg_mp3  # noqa: E402


class ConvertCommandTests(unittest.TestCase):
    def test_command_transcodes_to_192k_mp3_without_a_probe_step(self):
        command = convert_ogg_mp3.build_convert_command("ffmpeg.exe", "in.ogg", "out.mp3")
        self.assertEqual(command, ["ffmpeg.exe", "-y", "-i", "in.ogg", "-vn", "-b:a", "192k", "out.mp3"])

    def test_module_no_longer_depends_on_pydub(self):
        source = Path(convert_ogg_mp3.__file__).read_text(encoding="utf-8")
        self.assertNotIn("pydub", source)


if __name__ == "__main__":
    unittest.main()
