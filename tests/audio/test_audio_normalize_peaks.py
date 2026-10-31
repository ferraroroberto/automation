"""Sample peak must be requested from ffmpeg and parsed (no ffmpeg or audio needed)."""

import sys
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "audio"))

import audio_normalize  # noqa: E402

EBUR128_STDERR = """
  Integrated loudness:
    I:         -20.0 LUFS
    Threshold: -30.0 LUFS

  Loudness range:
    LRA:         5.0 LU

  Sample peak:
    Peak:      -4.2 dBFS

  True peak:
    Peak:      -3.9 dBFS
"""


class SamplePeakTests(unittest.TestCase):
    def test_parser_reads_both_peak_sections(self):
        analysis = audio_normalize._parse_ebur128_output(EBUR128_STDERR)
        self.assertEqual(analysis["sample_peak"], -4.2)
        self.assertEqual(analysis["true_peak"], -3.9)

    def test_analysis_asks_ffmpeg_for_sample_peak_as_well_as_true_peak(self):
        completed = SimpleNamespace(returncode=0, stderr=EBUR128_STDERR, stdout="")
        with mock.patch.object(audio_normalize.subprocess, "run", return_value=completed) as run:
            audio_normalize.analyze_audio_levels("in.wav")
        command = run.call_args.args[0]
        self.assertIn("ebur128=peak=true+sample", command)


if __name__ == "__main__":
    unittest.main()
