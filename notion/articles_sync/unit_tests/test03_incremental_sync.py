#!/usr/bin/env python3
"""
Unit tests for incremental sync change detection to identify hanging issues.
"""

import json
import time
import unittest
from unittest.mock import Mock, patch, MagicMock
import threading
from notion_articles_sync import NotionArticlesSync, RateLimiter


class TestIncrementalSync(unittest.TestCase):
    """Test cases for incremental sync functionality."""

    def setUp(self):
        """Set up test fixtures."""
        self.mock_config = {
            "notion": {
                "api_token": "test_token",
                "api_base": "https://api.notion.com/v1",
                "version": "2022-06-28",
                "source_database": "67fbcee66711465c852ebf97303787a3",
                "target_database": "163c4531dea649a7b2eaaf37dd95c68c"
            },
            "sync": {
                "polling_interval_seconds": 300,
                "batch_size": 100,
                "max_retries": 3,
                "backoff_seconds": 2.0
            },
            "threading": {
                "max_workers": 5,
                "requests_per_second": 3.0,
                "burst_size": 10,
                "parallel_fetching": True,
                "parallel_operations": True,
                "operation_batch_size": 10
            },
            "field_mapping": {
                "source_to_target": {
                    "rowid": "source rowid",
                    "title": "title"
                }
            }
        }

    @patch('notion_articles_sync.NotionArticlesSync.fetch_all_items')
    @patch('notion_articles_sync.NotionArticlesSync.api_call')
    def test_detect_changes_incremental_basic(self, mock_api_call, mock_fetch_all):
        """Test basic incremental sync change detection."""
        # Mock the fetch_all_items to return sample changed items
        mock_changed_items = [
            {
                "id": "page_1",
                "properties": {
                    "rowid": {"rich_text": [{"plain_text": "rowid_1"}]},
                    "exclude archive": {"checkbox": False}
                }
            },
            {
                "id": "page_2",
                "properties": {
                    "rowid": {"rich_text": [{"plain_text": "rowid_2"}]},
                    "exclude archive": {"checkbox": False}
                }
            }
        ]
        mock_fetch_all.return_value = mock_changed_items

        # Mock API calls for existence checks - first exists, second doesn't
        mock_api_call.side_effect = [
            {"results": [{"id": "existing_page"}]},  # page_1 exists
            {"results": []}  # page_2 doesn't exist
        ]

        # Create sync instance
        sync = NotionArticlesSync.__new__(NotionArticlesSync)
        sync.source_db = "source_db"
        sync.target_db = "target_db"
        sync.api_call = mock_api_call
        sync.get_last_sync_time = Mock(return_value=None)  # No previous sync

        # Test incremental detection
        new_items, updated_items, deleted_items = sync.detect_changes_incremental()

        # Should have detected 1 new item, 1 updated item
        self.assertEqual(len(new_items), 1)
        self.assertEqual(len(updated_items), 1)
        self.assertEqual(len(deleted_items), 0)

        # Verify API calls were made correctly
        self.assertEqual(mock_api_call.call_count, 2)

    @patch('notion_articles_sync.NotionArticlesSync.fetch_all_items')
    @patch('notion_articles_sync.NotionArticlesSync.api_call')
    def test_detect_changes_incremental_with_exclusions(self, mock_api_call, mock_fetch_all):
        """Test incremental sync with excluded items."""
        # Mock items - one excluded, one not
        mock_changed_items = [
            {
                "id": "page_1",
                "properties": {
                    "rowid": {"rich_text": [{"plain_text": "rowid_1"}]},
                    "exclude archive": {"checkbox": True}  # Excluded
                }
            },
            {
                "id": "page_2",
                "properties": {
                    "rowid": {"rich_text": [{"plain_text": "rowid_2"}]},
                    "exclude archive": {"checkbox": False}  # Not excluded
                }
            }
        ]
        mock_fetch_all.return_value = mock_changed_items

        # Mock API call for the non-excluded item
        mock_api_call.return_value = {"results": []}  # Doesn't exist

        # Create sync instance
        sync = NotionArticlesSync.__new__(NotionArticlesSync)
        sync.source_db = "source_db"
        sync.target_db = "target_db"
        sync.api_call = mock_api_call
        sync.get_last_sync_time = Mock(return_value=None)

        # Test incremental detection
        new_items, updated_items, deleted_items = sync.detect_changes_incremental()

        # Should have only processed the non-excluded item
        self.assertEqual(len(new_items), 1)
        self.assertEqual(len(updated_items), 0)
        self.assertEqual(len(deleted_items), 0)

        # Should have made only 1 API call (for the non-excluded item)
        self.assertEqual(mock_api_call.call_count, 1)

    @patch('notion_articles_sync.NotionArticlesSync.fetch_all_items')
    @patch('notion_articles_sync.NotionArticlesSync.api_call')
    def test_detect_changes_incremental_api_timeout_simulation(self, mock_api_call, mock_fetch_all):
        """Test incremental sync when API calls hang (simulate timeout)."""
        # Mock a smaller set of changed items for testing
        mock_changed_items = [
            {
                "id": "page_1",
                "properties": {
                    "rowid": {"rich_text": [{"plain_text": "rowid_1"}]},
                    "exclude archive": {"checkbox": False}
                }
            }
        ]
        mock_fetch_all.return_value = mock_changed_items

        # Mock API call that hangs/times out
        mock_api_call.return_value = None  # Simulate timeout/failure

        # Create sync instance
        sync = NotionArticlesSync.__new__(NotionArticlesSync)
        sync.source_db = "source_db"
        sync.target_db = "target_db"
        sync.api_call = mock_api_call
        sync.get_last_sync_time = Mock(return_value=None)

        # Test incremental detection - should handle API failures gracefully
        new_items, updated_items, deleted_items = sync.detect_changes_incremental()

        # Should have no changes due to API failure
        self.assertEqual(len(new_items), 0)
        self.assertEqual(len(updated_items), 0)
        self.assertEqual(len(deleted_items), 0)

    def test_large_dataset_processing_simulation(self):
        """Simulate processing a large dataset like the stuck scenario."""
        # Create a sync instance with mocked dependencies
        sync = NotionArticlesSync.__new__(NotionArticlesSync)
        sync.source_db = "source_db"
        sync.target_db = "target_db"
        sync.api_call_count = 0
        sync.api_call_lock = threading.Lock()

        # Mock rate limiter with reasonable settings
        sync.rate_limiter = RateLimiter(requests_per_second=3.0, burst_size=10)

        # Mock API call that succeeds but has some delay
        def mock_api_call(endpoint, method="GET", data=None, retry=0):
            time.sleep(0.01)  # Small delay to simulate network
            return {"results": []}  # Simulate item doesn't exist

        sync.api_call = mock_api_call

        # Create mock items similar to the stuck scenario (but much smaller for testing)
        mock_items = []
        for i in range(10):  # Small number for testing, original was 3170
            mock_items.append({
                "id": f"page_{i}",
                "properties": {
                    "rowid": {"rich_text": [{"plain_text": f"rowid_{i}"}]},
                    "exclude archive": {"checkbox": False}
                }
            })

        # Test the processing loop logic
        new_items = []
        updated_items = []
        start_time = time.monotonic()

        for item in mock_items:
            # This mimics the loop in detect_changes_incremental
            if sync.extract_value(item.get("properties", {}).get("exclude archive", {})):
                continue

            source_rowid = sync.normalize_rowid(
                sync.extract_value(item.get("properties", {}).get("rowid", {}))
            )

            if source_rowid:
                # This is the API call that might hang
                data = {
                    "filter": {
                        "property": "source rowid",
                        "rich_text": {"contains": source_rowid}
                    },
                    "page_size": 1
                }
                result = sync.api_call(f"databases/{sync.target_db}/query", method="POST", data=data)

                if result and result.get("results"):
                    updated_items.append((item, result["results"][0]))
                else:
                    new_items.append(item)

        elapsed = time.monotonic() - start_time

        # Should have processed all items
        self.assertEqual(len(new_items), 10)  # All items should be "new"
        self.assertEqual(len(updated_items), 0)

        # Should complete in reasonable time (much less than hanging)
        self.assertLess(elapsed, 5.0)  # Should complete within 5 seconds

    def test_extract_value_and_normalize_rowid(self):
        """Test the helper methods used in the processing loop."""
        sync = NotionArticlesSync.__new__(NotionArticlesSync)

        # Test extract_value with various property types
        test_cases = [
            ({"checkbox": False}, False),
            ({"rich_text": [{"plain_text": "test_text"}]}, "test_text"),
            ({"number": 42}, 42),
        ]

        for prop, expected in test_cases:
            result = sync.extract_value(prop)
            self.assertEqual(result, expected)

        # Test normalize_rowid
        test_rowids = [
            ("abc123def456", "abc123def456"),  # No dashes
            ("abc-123-def-456", "abc123def456"),  # With dashes
            ("", ""),  # Empty
            (None, None),  # None
        ]

        for input_rowid, expected in test_rowids:
            result = sync.normalize_rowid(input_rowid)
            self.assertEqual(result, expected)


if __name__ == '__main__':
    unittest.main(verbosity=2)
