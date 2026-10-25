"""Every check_ip HTTP call passes a timeout, and a timeout is handled by the caller."""

import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import pandas as pd
import requests

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "linkedin/check_ip"))

import imgur_client  # noqa: E402
import search_client  # noqa: E402


def _timeout_of(call) -> float:
    return call.kwargs.get("timeout") or 0


class ImgurClientTimeouts(unittest.TestCase):
    def setUp(self):
        self.client = imgur_client.ImgurClient("id", "token")

    def test_upload_passes_a_timeout_and_gives_up_after_timeouts(self):
        with tempfile.TemporaryDirectory() as folder:
            image = Path(folder) / "a.png"
            image.write_bytes(b"x")
            with mock.patch.object(imgur_client.requests, "post", side_effect=requests.exceptions.Timeout("stalled")) as post, \
                    mock.patch.object(imgur_client.time, "sleep"):
                self.assertIsNone(self.client.upload_image(image, retries=2))
        self.assertEqual(post.call_count, 2)
        self.assertGreater(_timeout_of(post.call_args), 0)

    def test_validate_credentials_passes_a_timeout_and_reports_false(self):
        with mock.patch.object(imgur_client.requests, "get", side_effect=requests.exceptions.Timeout("stalled")) as get:
            self.assertFalse(self.client.validate_credentials())
        self.assertGreater(_timeout_of(get.call_args), 0)


class SearchClientTimeouts(unittest.TestCase):
    def setUp(self):
        self.client = search_client.SearchClient("key", "https://serpapi.test/search")

    def test_search_passes_a_timeout_and_returns_no_data(self):
        with mock.patch.object(search_client.requests, "get", side_effect=requests.exceptions.Timeout("stalled")) as get:
            data, _ = self.client.search_google_lens("http://img", "a.png", pd.DataFrame(), "visual_matches")
        self.assertIsNone(data)
        self.assertGreater(_timeout_of(get.call_args), 0)

    def test_validate_api_key_passes_a_timeout_and_reports_false(self):
        with mock.patch.object(search_client.requests, "get", side_effect=requests.exceptions.Timeout("stalled")) as get:
            self.assertFalse(self.client.validate_api_key())
        self.assertGreater(_timeout_of(get.call_args), 0)


if __name__ == "__main__":
    unittest.main()
