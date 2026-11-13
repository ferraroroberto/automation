"""The config file's log_level applies unless --log-level is given."""

import argparse
import logging
import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tools" / "image"))

import illustrations_subscribers  # noqa: E402


class LogLevelTests(unittest.TestCase):
    def _args(self, log_level):
        return argparse.Namespace(
            subscribers=None, illustrations=None, source=None, archive=None, encode=None, log_level=log_level
        )

    def _config_file(self, tmp, level):
        path = Path(tmp) / "cfg.json"
        path.write_text('{"log_level": "%s"}' % level, encoding="utf-8")
        return str(path)

    def test_flag_defaults_to_none(self):
        with mock.patch.object(sys, "argv", ["illustrations_subscribers.py"]):
            self.assertIsNone(illustrations_subscribers.parse_arguments().log_level)

    def test_config_level_survives_when_flag_is_absent(self):
        import tempfile

        with tempfile.TemporaryDirectory() as tmp:
            config = illustrations_subscribers.load_config(
                self._config_file(tmp, "DEBUG"), self._args(None), logging.getLogger("t")
            )
        self.assertEqual(config["log_level"], "DEBUG")

    def test_flag_overrides_config(self):
        import tempfile

        with tempfile.TemporaryDirectory() as tmp:
            config = illustrations_subscribers.load_config(
                self._config_file(tmp, "DEBUG"), self._args("ERROR"), logging.getLogger("t")
            )
        self.assertEqual(config["log_level"], "ERROR")


if __name__ == "__main__":
    unittest.main()
