#!/usr/bin/env python3
"""
Tests for `_values_equivalent` and `items_are_different`.

These cover the cross-schema serialization quirks that caused the previous
incremental sync to flag every Notion-bumped item as 'changed':

  - 'created' field returned as `...Z` from one DB and `...+00:00` from the other
  - same multi_select option order varies between writes
  - source uses `select` while target uses `multi_select` for the same concept
  - empty values come back as None, "", or [] depending on property type
"""

import unittest

from _sync_fixtures import (
    prop_rich_text,
    prop_title,
    prop_formula_string,
    prop_select,
    prop_multi_select,
    prop_date,
    make_sync,
)


class TestValuesEquivalent(unittest.TestCase):

    def setUp(self):
        self.sync = make_sync()

    def test_identical_values(self):
        self.assertTrue(self.sync._values_equivalent("foo", "foo"))
        self.assertTrue(self.sync._values_equivalent(None, None))

    def test_empty_variants_are_equivalent(self):
        for empty in (None, "", []):
            for other in (None, "", []):
                self.assertTrue(
                    self.sync._values_equivalent(empty, other),
                    f"{empty!r} should equal {other!r}",
                )

    def test_iso_date_z_vs_offset(self):
        """The exact bug: source returns 'Z', target returns '+00:00'."""
        self.assertTrue(self.sync._values_equivalent(
            "2025-09-20T06:55:00.000Z",
            "2025-09-20T06:55:00.000+00:00",
        ))

    def test_iso_date_different_instants_differ(self):
        self.assertFalse(self.sync._values_equivalent(
            "2025-09-20T06:55:00.000Z",
            "2025-09-20T06:56:00.000+00:00",
        ))

    def test_select_vs_multi_select_single_value(self):
        self.assertTrue(self.sync._values_equivalent("ai", ["ai"]))
        self.assertTrue(self.sync._values_equivalent(["ai"], "ai"))

    def test_select_vs_multi_select_multiple_values_differ(self):
        self.assertFalse(self.sync._values_equivalent("ai", ["ai", "ml"]))

    def test_multi_select_order_independent(self):
        self.assertTrue(self.sync._values_equivalent(["a", "b"], ["b", "a"]))
        self.assertFalse(self.sync._values_equivalent(["a", "b"], ["a", "c"]))

    def test_unrelated_strings_differ(self):
        self.assertFalse(self.sync._values_equivalent("foo", "bar"))


class TestItemsAreDifferent(unittest.TestCase):
    """Confirms the comparator ignores last_edited_time and walks mapped fields."""

    def setUp(self):
        self.sync = make_sync()

    def _src(self, **overrides):
        props = {
            "rowid": prop_formula_string("aaaa"),
            "article": prop_title("Hello"),
            "topic": prop_select("ai"),
            "created": prop_date("2025-09-20T06:55:00.000Z"),
        }
        props.update(overrides)
        return {"id": "src", "last_edited_time": "2026-01-11T13:42:00.000Z", "properties": props}

    def _tgt(self, **overrides):
        props = {
            "source rowid": prop_rich_text("aaaa"),
            "article": prop_title("Hello"),
            "topic": prop_multi_select("ai"),
            "created": prop_date("2025-09-20T06:55:00.000+00:00"),
        }
        props.update(overrides)
        return {"id": "tgt", "last_edited_time": "2025-11-23T00:00:00.000Z", "properties": props}

    def test_identical_payload_with_schema_diffs_is_not_different(self):
        self.assertFalse(self.sync.items_are_different(self._src(), self._tgt()))

    def test_source_newer_but_same_content_is_not_different(self):
        """Even though source.last_edited_time >> target.last_edited_time."""
        self.assertFalse(self.sync.items_are_different(
            self._src(),
            self._tgt(),
        ))

    def test_real_field_change_is_detected(self):
        self.assertTrue(self.sync.items_are_different(
            self._src(article=prop_title("Changed")),
            self._tgt(),
        ))

    def test_topic_added_is_detected(self):
        self.assertTrue(self.sync.items_are_different(
            self._src(topic=prop_select("ai")),
            self._tgt(topic=prop_multi_select("ai", "ml")),
        ))


if __name__ == "__main__":
    unittest.main(verbosity=2)
