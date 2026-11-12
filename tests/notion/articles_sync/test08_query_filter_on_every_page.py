"""An incremental query repeats its last_edited_time filter on every page."""

import unittest
from datetime import datetime, timezone
from unittest.mock import Mock

from _sync_fixtures import make_sync


class TestQueryFilterOnEveryPage(unittest.TestCase):
    def _sent_body(self, start_cursor):
        sync = make_sync()
        sync.batch_size = 100
        sync.api_call = Mock(return_value={"results": []})
        sync.query_database("db", start_cursor, datetime(2026, 1, 1, tzinfo=timezone.utc))
        return sync.api_call.call_args.kwargs["data"]

    def test_first_page_is_filtered(self):
        self.assertIn("filter", self._sent_body(None))

    def test_continuation_page_keeps_the_filter_and_the_cursor(self):
        body = self._sent_body("cursor-1")
        self.assertEqual(body["start_cursor"], "cursor-1")
        self.assertEqual(body["filter"]["last_edited_time"]["after"], "2026-01-01T00:00:00+00:00")


if __name__ == "__main__":
    unittest.main()
