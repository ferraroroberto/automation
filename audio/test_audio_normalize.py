"""Regression tests for audio_normalize.normalize_audio (no ffmpeg, no real audio)."""

import os
import subprocess
import sys
import unittest
from types import SimpleNamespace
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import audio_normalize  # noqa: E402

LOUDNORM_JSON = """{
    "input_i": "-20.0",
    "input_tp": "-3.0",
    "input_lra": "5.0",
    "input_thresh": "-30.0",
    "target_offset": "0.5"
}"""


class NormalizeOutputExtensionTests(unittest.TestCase):
    def _run_normalize(self, input_path: str):
        calls = []

        def fake_run(cmd, **kwargs):
            calls.append(cmd)
            return SimpleNamespace(returncode=0, stdout="", stderr=LOUDNORM_JSON)

        with mock.patch.object(subprocess, "run", side_effect=fake_run):
            result = audio_normalize.normalize_audio(input_path, -16.0)
        return result, calls

    def test_output_is_m4a_for_every_picker_format(self):
        # The encode is always AAC in an MP4 container (-movflags +faststart),
        # which ffmpeg cannot write into .mp3/.flac/.ogg/.aiff/.wma containers.
        for ext in (".mp3", ".wav", ".flac", ".aac", ".ogg", ".m4a", ".wma", ".aiff"):
            with self.subTest(ext=ext):
                result, calls = self._run_normalize(os.path.join("clips", f"take{ext}"))
                self.assertTrue(result.endswith("_normalized_-16.0LUFS.m4a"), result)
                self.assertEqual(calls[-1][-1], result)


if __name__ == "__main__":
    unittest.main()
