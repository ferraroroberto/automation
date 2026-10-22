"""Regression tests for audio_extractor_core output naming (no ffmpeg, no media)."""

import subprocess
import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "audio"))

from audio_extractor_core import AudioExtractor  # noqa: E402


class DottedFileNameTests(unittest.TestCase):
    def test_outputs_of_a_dotted_name_do_not_collapse(self):
        # Path.with_suffix('.mp3') replaces the last dotted part of the *stem*
        # name, so "Movie.2020.1080p_audio_unified" lost "_audio_unified" and
        # every output overwrote the others under ffmpeg's -y.
        video = Path("clips") / "Movie.2020.1080p.mkv"
        extractor = AudioExtractor("clips")
        tracks = [{"index": 1}, {"index": 2}]
        written = []

        def fake_run(cmd, **kwargs):
            written.append(cmd[-1])
            return mock.Mock(returncode=0)

        with mock.patch.object(extractor, "get_audio_tracks", return_value=tracks), \
                mock.patch.object(subprocess, "run", side_effect=fake_run):
            extractor.process_video_file(video)

        names = sorted(Path(p).name for p in written)
        self.assertEqual(
            names,
            [
                "Movie.2020.1080p_audio_track1.mp3",
                "Movie.2020.1080p_audio_track2.mp3",
                "Movie.2020.1080p_audio_unified.mp3",
            ],
        )

    def test_single_track_unified_keeps_dotted_name(self):
        video = Path("clips") / "Movie.2020.1080p.mkv"
        extractor = AudioExtractor("clips")
        written = []

        def fake_run(cmd, **kwargs):
            written.append(cmd[-1])
            return mock.Mock(returncode=0)

        with mock.patch.object(extractor, "get_audio_tracks", return_value=[{"index": 1}]), \
                mock.patch.object(subprocess, "run", side_effect=fake_run):
            extractor.process_video_file(video)

        self.assertEqual([Path(p).name for p in written], ["Movie.2020.1080p_audio_unified.mp3"])


if __name__ == "__main__":
    unittest.main()
