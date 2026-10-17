#!/usr/bin/env python3
"""
A page query that fails after retries must abort the fetch, not truncate it.

`api_call` returns None once its retries are spent. `fetch_all_items` used to
`break` on that and hand back the partial list as if complete; in a full sync
every target row missing from a truncated source list was then archived as
deleted, and in the incremental loop a truncated target list made existing
rows look new and duplicated them.

The Notion client is faked at `api_call`, so no network is touched.
"""

import threading
import unittest
from unittest.mock import Mock

from notion_articles_sync import NotionArticlesSync


def _rich_text(value):
    return {"type": "rich_text", "rich_text": [{"plain_text": value, "text": {"content": value}}]}


def _page(**props):
    return {"properties": props, "last_edited_time": "2026-01-01T00:00:00.000Z"}


def _make_sync(pages_by_db):
    """NotionArticlesSync wired to a fake client.

    `pages_by_db` maps a database id to the sequence of `api_call` results for
    its successive page queries; None stands for a call that failed after retries.
    """
    sync = NotionArticlesSync.__new__(NotionArticlesSync)
    sync.source_db = "source_db"
    sync.target_db = "target_db"
    sync.batch_size = 100
    sync.debug = False
    sync.parallel_fetching = False
    sync.operation_batch_size = 10
    sync.field_mapping = {"rowid": "source rowid", "article": "article"}
    sync.api_call_count = 0
    sync.api_call_lock = threading.Lock()
    sync._schema_cache = {}
    sync.last_sync_time = None
    sync.save_sync_time = Mock()

    queues = {db: list(pages) for db, pages in pages_by_db.items()}

    def fake_api_call(endpoint, method="GET", data=None, retry=0):
        db = endpoint.split("/")[1]
        return queues[db].pop(0)

    sync.api_call = fake_api_call
    return sync


class TestFetchFailureAbortsSync(unittest.TestCase):

    def test_failed_page_raises_instead_of_returning_a_partial_list(self):
        sync = _make_sync({
            "source_db": [
                {"results": [_page(rowid=_rich_text("r1"))], "has_more": True, "next_cursor": "c1"},
                None,
            ],
        })
        with self.assertRaises(RuntimeError):
            sync.fetch_all_items("source_db")

    def test_complete_fetch_still_returns_every_page(self):
        sync = _make_sync({
            "source_db": [
                {"results": [_page(rowid=_rich_text("r1"))], "has_more": True, "next_cursor": "c1"},
                {"results": [_page(rowid=_rich_text("r2"))], "has_more": False},
            ],
        })
        self.assertEqual(len(sync.fetch_all_items("source_db")), 2)

    def test_truncated_source_does_not_archive_target_rows(self):
        # Source page 2 fails; the target row for "r2" lives on that lost page.
        sync = _make_sync({
            "source_db": [
                {"results": [_page(rowid=_rich_text("r1"))], "has_more": True, "next_cursor": "c1"},
                None,
            ],
            "target_db": [
                {
                    "results": [
                        _page(**{"source rowid": _rich_text("r1")}),
                        _page(**{"source rowid": _rich_text("r2")}),
                    ],
                    "has_more": False,
                },
            ],
        })
        sync.items_are_different = Mock(return_value=False)
        sync.process_batch_operations = Mock(return_value=(0, 0))

        self.assertFalse(sync.run_sync(force_full_sync=True))
        sync.process_batch_operations.assert_not_called()
        sync.save_sync_time.assert_not_called()


if __name__ == "__main__":
    unittest.main()
