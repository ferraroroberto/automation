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

import unittest
from unittest.mock import Mock

from _sync_fixtures import (
    prop_rich_text,
    prop_title,
    prop_formula_string,
    prop_select,
    prop_multi_select,
    prop_date,
    prop_checkbox,
    make_sync,
)


def _make_sync(field_mapping=None):
    sync = make_sync(field_mapping)
    sync.source_db = "source_db"
    sync.target_db = "target_db"
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
                "rowid": prop_formula_string("2740f91db10680138197ce92c074876d"),
                "article": prop_title("Virtual communication workshop ideas"),
                "topic": prop_select("personal development"),
                # source 'created' uses 'Z' suffix
                "created": prop_date("2025-09-20T06:55:00.000Z"),
                "exclude archive": prop_checkbox(False),
            },
        }
        target_item = {
            "id": "tgt_1",
            "last_edited_time": "2025-11-23T18:00:00.000Z",
            "properties": {
                "source rowid": prop_rich_text("2740f91db10680138197ce92c074876d"),
                "article": prop_title("Virtual communication workshop ideas"),
                # target stores topic as multi_select with the same single value
                "topic": prop_multi_select("personal development"),
                # target's 'created' came back from Notion with '+00:00' offset
                "created": prop_date("2025-09-20T06:55:00.000+00:00"),
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
                "rowid": prop_formula_string("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"),
                "article": prop_title("New title"),
                "topic": prop_select("ai"),
                "created": prop_date("2025-09-20T06:55:00.000Z"),
                "exclude archive": prop_checkbox(False),
            },
        }
        target_item = {
            "id": "tgt_2",
            "last_edited_time": "2025-11-23T18:00:00.000Z",
            "properties": {
                "source rowid": prop_rich_text("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"),
                "article": prop_title("Old title"),  # ← real diff
                "topic": prop_multi_select("ai"),
                "created": prop_date("2025-09-20T06:55:00.000+00:00"),
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
                "rowid": prop_formula_string("bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb"),
                "article": prop_title("Brand new"),
                "exclude archive": prop_checkbox(False),
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
                "rowid": prop_formula_string("cccccccccccccccccccccccccccccccc"),
                "article": prop_title("excluded"),
                "exclude archive": prop_checkbox(True),
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
