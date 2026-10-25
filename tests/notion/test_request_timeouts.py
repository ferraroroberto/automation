"""Notion helpers bound every request and route a timeout through their existing error handling."""

import sys
import unittest
from pathlib import Path
from unittest import mock

import requests

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "notion"))

import normalize_names  # noqa: E402
import normalize_url  # noqa: E402
import utils as notion_utils  # noqa: E402

TIMEOUT = requests.exceptions.Timeout("stalled")


def _timeout_of(call) -> float:
    return call.kwargs.get("timeout") or 0


class PaginatedQueryTimeout(unittest.TestCase):
    def test_post_passes_a_timeout_and_the_error_propagates(self):
        with mock.patch.object(notion_utils.requests, "post", side_effect=TIMEOUT) as post:
            with self.assertRaises(requests.exceptions.Timeout):
                notion_utils.paginated_database_query({}, "db")
        self.assertGreater(_timeout_of(post.call_args), 0)


class NormalizerPatchTimeouts(unittest.TestCase):
    def _normalizer(self, cls):
        # Skip __init__: it loads config, credentials and spaCy, none of which the PATCH needs.
        instance = cls.__new__(cls)
        instance.headers = {}
        return instance

    def test_name_update_passes_a_timeout_and_reports_failure(self):
        normalizer = self._normalizer(normalize_names.NotionNameNormalizer)
        with mock.patch.object(normalize_names.requests, "patch", side_effect=TIMEOUT) as patch:
            self.assertFalse(normalizer._update_page_property("p" * 32, "Name", "value"))
        self.assertGreater(_timeout_of(patch.call_args), 0)

    def test_url_update_passes_a_timeout_and_reports_failure(self):
        normalizer = self._normalizer(normalize_url.NotionURLNormalizer)
        with mock.patch.object(normalize_url.requests, "patch", side_effect=TIMEOUT) as patch:
            self.assertFalse(normalizer._update_page_url("p" * 32, "https://example.test"))
        self.assertGreater(_timeout_of(patch.call_args), 0)


if __name__ == "__main__":
    unittest.main()
