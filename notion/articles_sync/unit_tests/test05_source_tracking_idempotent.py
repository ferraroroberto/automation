#!/usr/bin/env python3
"""
Tests for the `update_source_tracking` no-op guard.

Without the guard, every successful update writes `target rowid` back to the
source — even when the source already tracks that archive id. That write
bumps source.last_edited_time, which makes the same item match the next
incremental filter, which causes another write, and so on. The guard breaks
the loop by skipping the PATCH when the value already matches.
"""

import threading
import unittest
from unittest.mock import Mock

from notion_articles_sync import NotionArticlesSync


def _make_sync():
    sync = NotionArticlesSync.__new__(NotionArticlesSync)
    sync.source_db = "source_db"
    sync.target_db = "target_db"
    sync.api_call_count = 0
    sync.api_call_lock = threading.Lock()
    return sync


class TestUpdateSourceTracking(unittest.TestCase):

    def test_skips_patch_when_value_already_matches(self):
        sync = _make_sync()
        sync.api_call = Mock()
        ok = sync.update_source_tracking(
            source_id="src_id",
            archive_id="abcd-1234",
            current_target_rowid="abcd1234",  # already normalized form, equal after dash strip
        )
        self.assertTrue(ok)
        sync.api_call.assert_not_called()

    def test_writes_patch_when_value_differs(self):
        sync = _make_sync()
        sync.api_call = Mock(return_value={"id": "src_id"})
        ok = sync.update_source_tracking(
            source_id="src_id",
            archive_id="abcd-1234",
            current_target_rowid="ffffffff",
        )
        self.assertTrue(ok)
        sync.api_call.assert_called_once()
        endpoint, kwargs = sync.api_call.call_args[0][0], sync.api_call.call_args[1]
        self.assertEqual(endpoint, "pages/src_id")
        self.assertEqual(kwargs.get("method"), "PATCH")
        self.assertIn("target rowid", kwargs["data"]["properties"])

    def test_writes_patch_when_no_current_value_known(self):
        sync = _make_sync()
        sync.api_call = Mock(return_value={"id": "src_id"})
        ok = sync.update_source_tracking(
            source_id="src_id",
            archive_id="abcd-1234",
            current_target_rowid=None,
        )
        self.assertTrue(ok)
        sync.api_call.assert_called_once()


if __name__ == "__main__":
    unittest.main(verbosity=2)
