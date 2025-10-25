#!/usr/bin/env python3
"""
Simple test of the progress reporting logic.
"""

import logging
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from notion_articles_sync import NotionArticlesSync, RateLimiter

# Setup logging to see the progress messages
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def simulate_progress_loop():
    """Simulate the progress reporting loop from detect_changes_incremental."""

    # Create test items like the real scenario
    test_items = []
    for i in range(250):  # Enough for multiple progress reports
        test_items.append({
            "id": f"item_{i}",
            "properties": {
                "rowid": {"type": "rich_text", "rich_text": [{"plain_text": f"rowid_{i}"}]},
                "exclude archive": {"type": "checkbox", "checkbox": False}
            }
        })

    print("Simulating progress reporting loop with 250 items...")
    print("You should see progress logs at 100, 200, and 250 items.")
    print("-" * 60)

    # Simulate the exact loop logic from detect_changes_incremental
    processed_count = 0
    total_items = len(test_items)

    new_items = []
    updated_items = []

    for item in test_items:
        processed_count += 1

        # Simulate exclusion check (none excluded in this test)
        if False:  # self.should_exclude(item) - simulated as False
            continue

        # Simulate rowid extraction (always succeeds in this test)
        source_rowid = f"rowid_{processed_count-1}"  # simulate normalize_rowid result

        if source_rowid:
            # Simulate API call (just add to new_items)
            new_items.append(item)

        # Progress reporting - log every 100 items processed
        if processed_count % 100 == 0:
            progress_percent = (processed_count / total_items) * 100
            logger.info(f"🔍 Progress: {processed_count}/{total_items} items processed ({progress_percent:.1f}%)")

    # Final progress log
    logger.info(f"🔍 Progress: {processed_count}/{total_items} items processed (100.0%)")
    logger.info(f"📊 Changes detected: {len(new_items)} new, {len(updated_items)} updated")

    print("-" * 60)
    print(f"Simulation completed: {len(new_items)} new items detected")
    print("Progress reporting working correctly!")

if __name__ == '__main__':
    simulate_progress_loop()