"""`imgur_check_rate` bounds its credits request and survives a stalled connection."""

import importlib.util
import unittest
from pathlib import Path
from unittest import mock

import requests

REPO_ROOT = Path(__file__).resolve().parents[2]


def _load_script():
    spec = importlib.util.spec_from_file_location("imgur_check_rate", REPO_ROOT / "image/imgur_check_rate.py")
    module = importlib.util.module_from_spec(spec)
    with mock.patch("dotenv.load_dotenv"):
        spec.loader.exec_module(module)
    return module


class ImgurCheckRateTimeout(unittest.TestCase):
    def setUp(self):
        self.script = _load_script()
        patcher = mock.patch.object(self.script, "ACCESS_TOKEN", "token")
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_credits_request_passes_a_timeout(self):
        response = mock.Mock(status_code=500, text="boom")
        with mock.patch.object(self.script.requests, "get", return_value=response) as get:
            self.script.check_imgur_rate_limit()
        self.assertGreater(get.call_args.kwargs.get("timeout", 0), 0)

    def test_timeout_is_logged_not_raised(self):
        with mock.patch.object(self.script.requests, "get", side_effect=requests.exceptions.Timeout("stalled")):
            with self.assertLogs(self.script.log, level="ERROR"):
                self.script.check_imgur_rate_limit()


if __name__ == "__main__":
    unittest.main()
