#!/usr/bin/env python3
"""
Unit tests for RateLimiter class to ensure it doesn't hang.
"""

import time
import threading
import unittest
from unittest.mock import patch
from notion_articles_sync import RateLimiter


class TestRateLimiter(unittest.TestCase):
    """Test cases for RateLimiter functionality."""

    def test_basic_rate_limiting(self):
        """Test basic rate limiting functionality."""
        limiter = RateLimiter(requests_per_second=2, burst_size=3)

        # Should be able to acquire tokens quickly
        start_time = time.monotonic()
        for i in range(3):
            self.assertTrue(limiter.acquire(timeout=1.0))
        elapsed = time.monotonic() - start_time

        # Should have used all burst tokens quickly
        self.assertLess(elapsed, 0.5)

        # Next acquire should wait for token regeneration
        start_time = time.monotonic()
        self.assertTrue(limiter.acquire(timeout=1.0))
        elapsed = time.monotonic() - start_time

        # Should have waited about 0.5 seconds (1/2 rate)
        self.assertGreaterEqual(elapsed, 0.4)
        self.assertLess(elapsed, 1.0)

    def test_timeout_behavior(self):
        """Test that acquire times out properly."""
        limiter = RateLimiter(requests_per_second=0.1, burst_size=1)

        # Use the burst token
        self.assertTrue(limiter.acquire(timeout=1.0))

        # Next acquire should timeout since rate is very low
        start_time = time.monotonic()
        result = limiter.acquire(timeout=0.1)
        elapsed = time.monotonic() - start_time

        self.assertFalse(result)  # Should timeout
        self.assertGreaterEqual(elapsed, 0.09)
        self.assertLess(elapsed, 0.15)

    def test_thread_safety(self):
        """Test that rate limiter works correctly with multiple threads."""
        limiter = RateLimiter(requests_per_second=10, burst_size=5)
        results = []
        errors = []

        def worker(worker_id):
            try:
                for i in range(3):
                    if limiter.acquire(timeout=1.0):
                        results.append(f"worker_{worker_id}_acquire_{i}")
                        time.sleep(0.01)  # Small delay to simulate work
                    else:
                        results.append(f"worker_{worker_id}_timeout_{i}")
            except Exception as e:
                errors.append(f"worker_{worker_id}_error: {e}")

        # Start multiple threads
        threads = []
        for i in range(5):
            t = threading.Thread(target=worker, args=(i,))
            threads.append(t)
            t.start()

        # Wait for all threads
        for t in threads:
            t.join()

        # Should have no errors
        self.assertEqual(len(errors), 0)

        # Should have acquired tokens (exact count depends on timing)
        acquire_count = len([r for r in results if 'acquire' in r])
        self.assertGreater(acquire_count, 0)

    def test_no_deadlock_under_load(self):
        """Test that rate limiter doesn't deadlock under heavy load.

        Sized so the total work (10 threads × 5 acquires = 50 tokens at 20 req/s
        = ~2.5s) finishes well inside the test timeout. The previous sizing
        (1 req/s, 50 tokens = 50s required, 10s allowed) reported as a deadlock
        when it was actually just throttled.
        """
        limiter = RateLimiter(requests_per_second=20, burst_size=5)

        # Simulate heavy concurrent load
        results = []
        errors = []

        def stress_worker(worker_id):
            try:
                for i in range(5):
                    start = time.monotonic()
                    acquired = limiter.acquire(timeout=2.0)  # Reasonable timeout
                    elapsed = time.monotonic() - start

                    if acquired:
                        results.append(f"worker_{worker_id}_success_{i}_{elapsed:.3f}s")
                    else:
                        results.append(f"worker_{worker_id}_timeout_{i}_{elapsed:.3f}s")

                    # Small delay between attempts
                    time.sleep(0.001)
            except Exception as e:
                errors.append(f"worker_{worker_id}_error: {e}")

        # Start many threads to stress test
        threads = []
        for i in range(10):
            t = threading.Thread(target=stress_worker, args=(i,))
            threads.append(t)
            t.start()

        # Wait for all with timeout
        start_wait = time.monotonic()
        for t in threads:
            remaining_time = max(0, 10 - (time.monotonic() - start_wait))  # 10s total timeout
            t.join(timeout=remaining_time)

        # Check for deadlocks (threads that didn't finish)
        active_threads = [t for t in threads if t.is_alive()]
        self.assertEqual(len(active_threads), 0, f"Deadlock detected: {len(active_threads)} threads still running")

        # Should have no errors
        self.assertEqual(len(errors), 0, f"Errors occurred: {errors}")

    def test_token_regeneration(self):
        """Test that tokens regenerate over time."""
        limiter = RateLimiter(requests_per_second=2, burst_size=2)

        # Use all tokens
        self.assertTrue(limiter.acquire(timeout=1.0))
        self.assertTrue(limiter.acquire(timeout=1.0))

        # Next should wait
        start = time.monotonic()
        self.assertTrue(limiter.acquire(timeout=1.0))
        elapsed = time.monotonic() - start

        # Should wait about 0.5 seconds for token regeneration
        self.assertGreaterEqual(elapsed, 0.4)
        self.assertLess(elapsed, 0.7)


if __name__ == '__main__':
    unittest.main(verbosity=2)
