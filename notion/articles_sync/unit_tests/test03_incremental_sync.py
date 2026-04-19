#!/usr/bin/env python3
"""
Tests for the incremental upsert path.

These cover the regression that motivated the upsert rewrite: Notion bumps
`last_edited_time` on non-content events (schema edits, formula recomputes,
write-back of `target rowid`), and the previous incremental loop trusted that
timestamp blindly — re-processing the entire database every run.

Specifically validates:
  - candidates whose mapped fields are semantically identical are skipped
  - candidates with real field differences are enqueued as updates
  - candidates without a target page are enqueued as new
  - excluded items are dropped before any target lookup
"""

import threading
import unittest
from unittest.mock import Mock

from notion_articles_sync import NotionArticlesSync


def _rich_text(value):
    return {"type": "rich_text", "rich_text": [{"plain_text": value, "text": {"content": value}}]}


def _title(value):
    return {"type": "title", "title": [{"plain_text": value, "text": {"content": value}}]}


def _formula_string(value):
    return {"type": "formula", "formula": {"type": "string", "string": value}}


def _select(name):
    return {"type": "select", "select": {"name": name}}


def _multi_select(*names):
    return {"type": "multi_select", "multi_select": [{"name": n} for n in names]}


def _date(start):
    return {"type": "date", "date": {"start": start}}


def _checkbox(value):
    return {"type": "checkbox", "checkbox": value}


def _make_sync(field_mapping=None):
    """Build a NotionArticlesSync without running __init__ side effects."""
    sync = NotionArticlesSync.__new__(NotionArticlesSync)
    sync.source_db = "source_db"
    sync.target_db = "target_db"
    sync.field_mapping = field_mapping or {
        "article": "article",
        "topic": "topic",
        "rowid": "source rowid",
        "created": "created",
    }
    sync.api_call_count = 0
    sync.api_call_lock = threading.Lock()
    sync._schema_cache = {}
    return sync


class TestIncrementalUpsert(unittest.TestCase):

    def test_skips_items_whose_fields_are_semantically_identical(self):
        """The bug: a Jan-11 mass time-bump made every item match the filter.
        After the fix, candidates whose mapped values match the target are skipped."""
        sync = _make_sync()
        sync.get_last_sync_time = Mock(return_value=None)
        # Bypass full-sync fallback by stubbing detect_changes_full_sync if called.
        sync.detect_changes_full_sync = Mock(return_value=([], [], []))
        sync.get_last_sync_time = Mock(return_value=Mock())  # truthy → incremental path

        source_item = {
            "id": "src_1",
            "last_edited_time": "2026-01-11T13:42:00.000Z",  # bumped, but content unchanged
            "properties": {
                "rowid": _formula_string("2740f91db10680138197ce92c074876d"),
                "article": _title("Virtual communication workshop ideas"),
                "topic": _select("personal development"),
                # source 'created' uses 'Z' suffix
                "created": _date("2025-09-20T06:55:00.000Z"),
                "exclude archive": _checkbox(False),
            },
        }
        target_item = {
            "id": "tgt_1",
            "last_edited_time": "2025-11-23T18:00:00.000Z",
            "properties": {
                "source rowid": _rich_text("2740f91db10680138197ce92c074876d"),
                "article": _title("Virtual communication workshop ideas"),
                # target stores topic as multi_select with the same single value
                "topic": _multi_select("personal development"),
                # target's 'created' came back from Notion with '+00:00' offset
                "created": _date("2025-09-20T06:55:00.000+00:00"),
            },
        }

        sync.fetch_all_items = Mock(side_effect=lambda db, **kw:
            [source_item] if db == sync.source_db else [target_item])

        new_items, updated_items, deleted_items = sync.detect_changes_incremental()

        self.assertEqual(new_items, [])
        self.assertEqual(updated_items, [],
                         "Item with no real diff should be skipped, not re-synced")
        self.assertEqual(deleted_items, [])

    def test_enqueues_real_field_changes(self):
        sync = _make_sync()
        sync.get_last_sync_time = Mock(return_value=Mock())

        source_item = {
            "id": "src_2",
            "last_edited_time": "2026-04-01T00:00:00.000Z",
            "properties": {
                "rowid": _formula_string("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"),
                "article": _title("New title"),
                "topic": _select("ai"),
                "created": _date("2025-09-20T06:55:00.000Z"),
                "exclude archive": _checkbox(False),
            },
        }
        target_item = {
            "id": "tgt_2",
            "last_edited_time": "2025-11-23T18:00:00.000Z",
            "properties": {
                "source rowid": _rich_text("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"),
                "article": _title("Old title"),  # ← real diff
                "topic": _multi_select("ai"),
                "created": _date("2025-09-20T06:55:00.000+00:00"),
            },
        }

        sync.fetch_all_items = Mock(side_effect=lambda db, **kw:
            [source_item] if db == sync.source_db else [target_item])

        new_items, updated_items, _ = sync.detect_changes_incremental()
        self.assertEqual(new_items, [])
        self.assertEqual(len(updated_items), 1)
        self.assertIs(updated_items[0][0], source_item)
        self.assertIs(updated_items[0][1], target_item)

    def test_enqueues_new_when_no_target(self):
        sync = _make_sync()
        sync.get_last_sync_time = Mock(return_value=Mock())

        source_item = {
            "id": "src_3",
            "last_edited_time": "2026-04-10T00:00:00.000Z",
            "properties": {
                "rowid": _formula_string("bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb"),
                "article": _title("Brand new"),
                "exclude archive": _checkbox(False),
            },
        }
        sync.fetch_all_items = Mock(side_effect=lambda db, **kw:
            [source_item] if db == sync.source_db else [])

        new_items, updated_items, _ = sync.detect_changes_incremental()
        self.assertEqual(updated_items, [])
        self.assertEqual(len(new_items), 1)

    def test_skips_excluded_items_without_target_lookup(self):
        sync = _make_sync()
        sync.get_last_sync_time = Mock(return_value=Mock())

        source_item = {
            "id": "src_4",
            "last_edited_time": "2026-04-10T00:00:00.000Z",
            "properties": {
                "rowid": _formula_string("cccccccccccccccccccccccccccccccc"),
                "article": _title("excluded"),
                "exclude archive": _checkbox(True),
            },
        }
        sync.fetch_all_items = Mock(side_effect=lambda db, **kw:
            [source_item] if db == sync.source_db else [])
        sync.api_call = Mock()

        new_items, updated_items, _ = sync.detect_changes_incremental()
        self.assertEqual(new_items, [])
        self.assertEqual(updated_items, [])
        sync.api_call.assert_not_called()


if __name__ == "__main__":
    unittest.main(verbosity=2)
