#!/usr/bin/env python3
"""
Unit tests for API call functionality to identify hanging issues.
"""

import json
import time
import unittest
from unittest.mock import Mock, patch, MagicMock
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from notion_articles_sync import NotionArticlesSync, RateLimiter


class TestApiCalls(unittest.TestCase):
    """Test cases for API call functionality."""

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

    @patch('notion_articles_sync.requests.get')
    def test_api_call_success(self, mock_get):
        """Test successful API call."""
        mock_response = Mock()
        mock_response.raise_for_status.return_value = None
        mock_response.json.return_value = {"test": "data"}
        mock_get.return_value = mock_response

        sync = NotionArticlesSync.__new__(NotionArticlesSync)
        sync.api_base = "https://api.notion.com/v1"
        sync.headers = {"Authorization": "Bearer test_token"}
        sync.rate_limiter = RateLimiter(10, 10)  # High rate for testing
        sync.max_retries = 3
        sync.api_call_count = 0
        sync.api_call_lock = threading.Lock()

        result = sync.api_call("test/endpoint")

        self.assertEqual(result, {"test": "data"})
        mock_get.assert_called_once()

    @patch('notion_articles_sync.requests.get')
    def test_api_call_retry_on_failure(self, mock_get):
        """Test API call retry logic."""
        # First two calls fail, third succeeds
        mock_response_fail = Mock()
        mock_response_fail.raise_for_status.side_effect = Exception("API Error")

        mock_response_success = Mock()
        mock_response_success.raise_for_status.return_value = None
        mock_response_success.json.return_value = {"success": True}

        mock_get.side_effect = [mock_response_fail, mock_response_fail, mock_response_success]

        sync = NotionArticlesSync.__new__(NotionArticlesSync)
        sync.api_base = "https://api.notion.com/v1"
        sync.headers = {"Authorization": "Bearer test_token"}
        sync.rate_limiter = RateLimiter(10, 10)
        sync.max_retries = 3
        sync.api_call_count = 0
        sync.api_call_lock = threading.Lock()

        result = sync.api_call("test/endpoint")

        self.assertEqual(result, {"success": True})
        self.assertEqual(mock_get.call_count, 3)

    @patch('notion_articles_sync.requests.get')
    @patch('notion_articles_sync.time.sleep')
    def test_api_call_max_retries_exceeded(self, mock_sleep, mock_get):
        """Test API call when max retries exceeded."""
        mock_response = Mock()
        mock_response.raise_for_status.side_effect = Exception("Persistent API Error")
        mock_get.return_value = mock_response

        sync = NotionArticlesSync.__new__(NotionArticlesSync)
        sync.api_base = "https://api.notion.com/v1"
        sync.headers = {"Authorization": "Bearer test_token"}
        sync.rate_limiter = RateLimiter(10, 10)
        sync.max_retries = 2  # Lower for testing
        sync.api_call_count = 0
        sync.api_call_lock = threading.Lock()

        result = sync.api_call("test/endpoint")

        self.assertIsNone(result)
        self.assertEqual(mock_get.call_count, 3)  # Initial + 2 retries

    def test_rate_limiter_integration(self):
        """Test that API calls properly integrate with rate limiter."""
        sync = NotionArticlesSync.__new__(NotionArticlesSync)
        sync.api_base = "https://api.notion.com/v1"
        sync.headers = {"Authorization": "Bearer test_token"}
        sync.rate_limiter = RateLimiter(1, 1)  # Very slow rate limiter
        sync.max_retries = 1
        sync.api_call_count = 0
        sync.api_call_lock = threading.Lock()

        # Mock successful response
        with patch('notion_articles_sync.requests.get') as mock_get:
            mock_response = Mock()
            mock_response.raise_for_status.return_value = None
            mock_response.json.return_value = {"test": "data"}
            mock_get.return_value = mock_response

            start_time = time.monotonic()

            # Make two rapid calls
            result1 = sync.api_call("test/endpoint1")
            result2 = sync.api_call("test/endpoint2")

            elapsed = time.monotonic() - start_time

            # Both should succeed
            self.assertIsNotNone(result1)
            self.assertIsNotNone(result2)

            # Should have taken at least 1 second due to rate limiting
            self.assertGreaterEqual(elapsed, 0.9)

    def test_concurrent_api_calls_no_deadlock(self):
        """Test concurrent API calls don't cause deadlocks."""
        sync = NotionArticlesSync.__new__(NotionArticlesSync)
        sync.api_base = "https://api.notion.com/v1"
        sync.headers = {"Authorization": "Bearer test_token"}
        sync.rate_limiter = RateLimiter(2, 3)  # Moderate rate limiting
        sync.max_retries = 1
        sync.api_call_count = 0
        sync.api_call_lock = threading.Lock()

        results = []
        errors = []

        def api_call_worker(call_id):
            try:
                with patch('notion_articles_sync.requests.get') as mock_get:
                    mock_response = Mock()
                    mock_response.raise_for_status.return_value = None
                    mock_response.json.return_value = {"call_id": call_id}
                    mock_get.return_value = mock_response

                    result = sync.api_call(f"test/endpoint/{call_id}")
                    results.append((call_id, result))
            except Exception as e:
                errors.append((call_id, str(e)))

        # Start concurrent API calls
        threads = []
        for i in range(5):
            t = threading.Thread(target=api_call_worker, args=(i,))
            threads.append(t)
            t.start()

        # Wait for completion with timeout
        start_wait = time.monotonic()
        for t in threads:
            remaining = max(0, 10 - (time.monotonic() - start_wait))
            t.join(timeout=remaining)

        # Check for deadlocks
        active_threads = [t for t in threads if t.is_alive()]
        self.assertEqual(len(active_threads), 0, f"Deadlock detected: {len(active_threads)} threads still running")

        # Should have no errors
        self.assertEqual(len(errors), 0, f"API call errors: {errors}")

        # Should have results for all calls
        self.assertEqual(len(results), 5)

    @patch('notion_articles_sync.requests.post')
    def test_database_query_call(self, mock_post):
        """Test the specific database query call that might be hanging."""
        mock_response = Mock()
        mock_response.raise_for_status.return_value = None
        mock_response.json.return_value = {
            "results": [{"id": "test_page", "properties": {"test": "value"}}],
            "has_more": False,
            "next_cursor": None
        }
        mock_post.return_value = mock_response

        sync = NotionArticlesSync.__new__(NotionArticlesSync)
        sync.api_base = "https://api.notion.com/v1"
        sync.headers = {"Authorization": "Bearer test_token"}
        sync.rate_limiter = RateLimiter(10, 10)
        sync.max_retries = 3
        sync.api_call_count = 0
        sync.api_call_lock = threading.Lock()
        sync.target_db = "test_db_id"

        # Test the exact call that happens in detect_changes_incremental
        data = {
            "filter": {
                "property": "source rowid",
                "rich_text": {"contains": "test_rowid"}
            },
            "page_size": 1
        }

        result = sync.api_call(f"databases/{sync.target_db}/query", method="POST", data=data)

        self.assertIsNotNone(result)
        self.assertIn("results", result)
        mock_post.assert_called_once()


if __name__ == '__main__':
    unittest.main(verbosity=2)
