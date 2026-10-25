"""Change detection for the articles sync.

`ChangeDetectionMixin` compares the source and target databases (full or
incremental) and classifies items as new, updated or deleted. It relies on the
`NotionArticlesSync` state, `fetch_all_items`, `get_last_sync_time` and the
`PropertyMappingMixin` helpers.
"""

from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional, Tuple

from sync_logger import logger


class ChangeDetectionMixin:
    """Index both databases by rowid and diff them."""

    def should_exclude(self, item: Dict[str, Any]) -> bool:
        """Check if item should be excluded from sync.
        
        Args:
            item: Notion item to check
            
        Returns:
            True if item should be excluded, False otherwise
        """
        exclude = self.extract_value(
            item.get("properties", {}).get("exclude archive", {})
        )
        return bool(exclude)
    
    @staticmethod
    def _normalize_iso_date(value: str) -> Optional[str]:
        """Return a canonical UTC ISO string for an ISO datetime, or None if unparseable.

        Notion serializes the same instant in different ways depending on the property
        type and how the value was last written (e.g. '2025-09-20T06:55:00.000Z' vs
        '2025-09-20T06:55:00.000+00:00'). Normalize so equality comparisons survive
        this round-trip.
        """
        try:
            dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
            return dt.astimezone(timezone.utc).isoformat()
        except (ValueError, TypeError, AttributeError):
            return None

    def _values_equivalent(self, source_value: Any, target_value: Any) -> bool:
        """Semantic equality for two extracted property values.

        Handles the cases where the source and target schemas legitimately store the
        same data in different shapes:
          - None / "" / [] all count as empty
          - ISO date strings differing only in 'Z' vs '+00:00' offset
          - select ↔ multi_select (e.g. "x" vs ["x"])
          - multi_select with the same elements in different order
        """
        if source_value == target_value:
            return True
        EMPTY = (None, "", [])
        if source_value in EMPTY and target_value in EMPTY:
            return True
        if isinstance(source_value, str) and isinstance(target_value, str):
            sn = self._normalize_iso_date(source_value)
            tn = self._normalize_iso_date(target_value)
            if sn is not None and tn is not None and sn == tn:
                return True
        if isinstance(source_value, str) and isinstance(target_value, list):
            return len(target_value) == 1 and target_value[0] == source_value
        if isinstance(target_value, str) and isinstance(source_value, list):
            return len(source_value) == 1 and source_value[0] == target_value
        if isinstance(source_value, list) and isinstance(target_value, list):
            try:
                return sorted(source_value) == sorted(target_value)
            except TypeError:
                return False
        return False

    def items_are_different(self, source_item: Dict[str, Any], target_item: Dict[str, Any]) -> bool:
        """True iff any mapped field semantically differs between source and target.

        Deliberately ignores last_edited_time: Notion bumps it for many reasons that
        do not change the fields we sync (schema edits, formula recomputes, our own
        write-back to `target rowid`). Trusting it produces 100% false positives
        after any DB-wide touch.
        """
        src_props = source_item.get("properties", {})
        tgt_props = target_item.get("properties", {})
        for src_field, tgt_field in self.field_mapping.items():
            sv = self.extract_value(src_props.get(src_field, {}))
            tv = self.extract_value(tgt_props.get(tgt_field, {}))
            if src_field == "rowid":
                sv = self.normalize_rowid(sv) if sv else sv
                tv = self.normalize_rowid(tv) if tv else tv
            if not self._values_equivalent(sv, tv):
                return True
        return False
    
    def _index_source_by_rowid(self, source_items: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
        """Index (non-excluded) source items by their 'rowid' property.

        Shared by detect_changes_full_sync() and detect_changes_incremental()
        (audit issue #67).
        """
        source_by_rowid: Dict[str, Dict[str, Any]] = {}
        for item in source_items:
            if self.should_exclude(item):
                continue
            rowid = self.normalize_rowid(
                self.extract_value(item.get("properties", {}).get("rowid", {}))
            )
            if rowid:
                source_by_rowid[rowid] = item
        return source_by_rowid

    def _index_target_by_source_rowid(self, target_items: List[Dict[str, Any]],
                                       warn_on_duplicate: bool = False) -> Dict[str, Dict[str, Any]]:
        """Index target items by their 'source rowid' property, keeping the
        most recently edited item when the same source rowid appears more
        than once.

        Shared by detect_changes_full_sync() and detect_changes_incremental()
        (audit issue #67).
        """
        target_by_source_rowid: Dict[str, Dict[str, Any]] = {}
        for item in target_items:
            source_rowid = self.normalize_rowid(
                self.extract_value(item.get("properties", {}).get("source rowid", {}))
            )
            if not source_rowid:
                continue
            existing = target_by_source_rowid.get(source_rowid)
            if existing is None:
                target_by_source_rowid[source_rowid] = item
                continue
            existing_time = existing.get("last_edited_time", "")
            new_time = item.get("last_edited_time", "")
            if new_time > existing_time:
                if warn_on_duplicate:
                    logger.warning(f"⚠️ Found duplicate for {source_rowid}, keeping newer")
                target_by_source_rowid[source_rowid] = item
        return target_by_source_rowid

    def _diff_source_and_target(
        self,
        source_by_rowid: Dict[str, Dict[str, Any]],
        target_by_source_rowid: Dict[str, Dict[str, Any]],
        detect_deletes: bool,
    ) -> Tuple[List[Dict[str, Any]], List[Tuple[Dict[str, Any], Dict[str, Any]]], List[Dict[str, Any]], int]:
        """Classify indexed source items against an indexed target as
        new/updated/unchanged, and (when detect_deletes) target items with no
        matching source as deleted.

        Shared diff step between detect_changes_full_sync() and
        detect_changes_incremental() (audit issue #67 — the two used to
        duplicate ~15-20 lines each of this same rowid -> target_by_rowid
        lookup + "keep newer of duplicate" classification).

        Returns:
            Tuple of (new_items, updated_items, deleted_items, unchanged_count)
        """
        new_items: List[Dict[str, Any]] = []
        updated_items: List[Tuple[Dict[str, Any], Dict[str, Any]]] = []
        deleted_items: List[Dict[str, Any]] = []
        unchanged_count = 0

        for rowid, source_item in source_by_rowid.items():
            target_item = target_by_source_rowid.get(rowid)
            if target_item is None:
                # Item only in source - needs to be created
                new_items.append(source_item)
            elif self.items_are_different(source_item, target_item):
                updated_items.append((source_item, target_item))
            else:
                unchanged_count += 1

        if detect_deletes:
            for source_rowid, target_item in target_by_source_rowid.items():
                if source_rowid not in source_by_rowid and not target_item.get("archived", False):
                    # Item only in target - has been deleted from source
                    deleted_items.append(target_item)

        return new_items, updated_items, deleted_items, unchanged_count

    def detect_changes_full_sync(self) -> Tuple[List[Dict[str, Any]], List[Tuple[Dict[str, Any], Dict[str, Any]]], List[Dict[str, Any]]]:
        """Detect changes using full database comparison.

        Returns:
            Tuple of (new_items, updated_items, deleted_items)
        """
        logger.info("🔍 Performing full sync - loading both databases...")

        # Reset API call counter for progress tracking
        with self.api_call_lock:
            self.api_call_count = 0

        # Load databases (both databases concurrently if enabled)
        if self.parallel_fetching:
            logger.info("📥 Loading databases in parallel...")

            with ThreadPoolExecutor(max_workers=2) as executor:
                source_future = executor.submit(self.fetch_all_items,
                                              self.source_db, None, True)
                target_future = executor.submit(self.fetch_all_items,
                                              self.target_db, None, True)

                source_items = source_future.result()
                target_items = target_future.result()
        else:
            # Sequential loading
            logger.info("📥 Loading source database...")
            source_items = self.fetch_all_items(self.source_db, show_progress=True)
            logger.info(f"📊 Loaded {len(source_items)} source items")

            logger.info("📥 Loading target database...")
            target_items = self.fetch_all_items(self.target_db, show_progress=True)
            logger.info(f"📊 Loaded {len(target_items)} target items")

        logger.info(f"📊 Total API calls for loading: {self.api_call_count}")

        # Build lookup maps
        logger.info("🔗 Building relationship maps...")
        source_by_rowid = self._index_source_by_rowid(source_items)
        target_by_source_rowid = self._index_target_by_source_rowid(target_items, warn_on_duplicate=True)

        logger.info("🔍 Analyzing changes...")
        new_items, updated_items, deleted_items, unchanged_count = self._diff_source_and_target(
            source_by_rowid, target_by_source_rowid, detect_deletes=True
        )

        logger.info(f"📊 Analysis complete:")
        logger.info(f"   ✨ New items: {len(new_items)}")
        logger.info(f"   🔄 Updated items: {len(updated_items)}")
        logger.info(f"   🗑️  Deleted items: {len(deleted_items)}")
        logger.info(f"   ✓ Unchanged items: {unchanged_count}")

        return new_items, updated_items, deleted_items

    def detect_changes_incremental(self) -> Tuple[List[Dict[str, Any]], List[Tuple[Dict[str, Any], Dict[str, Any]]], List[Dict[str, Any]]]:
        """Detect changes using incremental sync based on last sync time.

        Returns:
            Tuple of (new_items, updated_items, deleted_items)
        """
        logger.info("🔍 Detecting changes (incremental)...")

        # Get last sync time
        last_sync_time = self.get_last_sync_time()

        if last_sync_time:
            logger.info(f"⚡ Incremental sync - checking changes since {last_sync_time.isoformat()}")
        else:
            logger.info("📋 No previous sync found - performing full sync")
            return self.detect_changes_full_sync()

        # Fetch only changed items from source
        logger.info("📥 Fetching changed items from source...")
        source_items = self.fetch_all_items(self.source_db, filter_after=last_sync_time, show_progress=True)
        logger.info(f"📊 Found {len(source_items)} changed items")

        # Bulk-fetch the target once and index by 'source rowid' for O(1) lookup.
        # Previous code issued one filter query per source item (3500+ × ~1.7s ≈ 100 min).
        # Bulk fetch is ~35 pages (~90s) regardless of how many source items changed.
        logger.info("📥 Loading target database for lookup index...")
        target_items = self.fetch_all_items(self.target_db, show_progress=True)
        logger.info(f"📊 Loaded {len(target_items)} target items")

        source_by_rowid = self._index_source_by_rowid(source_items)
        target_by_source_rowid = self._index_target_by_source_rowid(target_items, warn_on_duplicate=False)

        new_items, updated_items, _deleted_items, unchanged_count = self._diff_source_and_target(
            source_by_rowid, target_by_source_rowid, detect_deletes=False
        )

        logger.info(f"📊 Changes detected: {len(new_items)} new, {len(updated_items)} updated, {unchanged_count} unchanged (skipped)")

        return new_items, updated_items, []
