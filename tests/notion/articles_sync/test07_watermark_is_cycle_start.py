#!/usr/bin/env python3
"""
The incremental watermark is the moment the cycle read the source.

Stamping it when the cycle ends skipped any article edited while a long cycle
ran: its `last_edited_time` fell before the saved watermark, so no later
incremental sync saw it until a full sync.
"""

import unittest
from datetime import datetime, timedelta, timezone
from unittest.mock import Mock, patch

from _sync_fixtures import make_sync
import notion_articles_sync

CYCLE_START = datetime(2026, 1, 1, 12, 0, 0, tzinfo=timezone.utc)


class _Clock:
    """`datetime` stand-in whose time only moves when the test advances it."""

    def __init__(self):
        self.elapsed = timedelta(0)

    def now(self, tz=None):
        return CYCLE_START + self.elapsed


class TestWatermarkIsCycleStart(unittest.TestCase):

    def test_saved_watermark_is_taken_before_the_source_fetch(self):
        clock = _Clock()

        def slow_fetch():
            clock.elapsed += timedelta(minutes=10)
            return [{"id": "n1"}], [], []

        sync = make_sync()
        sync.last_sync_time = None
        sync.save_sync_time = Mock()
        sync.operation_batch_size = 10
        sync.detect_changes_incremental = slow_fetch
        sync.process_batch_operations = Mock(return_value=(1, 0))

        with patch.object(notion_articles_sync, "datetime", clock):
            self.assertTrue(sync.run_sync())

        self.assertEqual(sync.last_sync_time, CYCLE_START)
        sync.save_sync_time.assert_called_once()


if __name__ == "__main__":
    unittest.main()
