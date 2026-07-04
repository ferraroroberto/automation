#!/usr/bin/env python3
"""
Notion Articles Sync - Syncs articles database to archive with full sync capabilities

Watches a source Notion database for changes and syncs them to a target archive database.
Handles create, update, and delete operations with efficient time-based filtering.
Supports parallel operations with configurable threading and rate limiting.

Usage:
    python notion_articles_sync.py
    python notion_articles_sync.py --config custom_config.json
    python notion_articles_sync.py --once
    python notion_articles_sync.py --full-sync
    python notion_articles_sync.py --debug
"""

import json
import os
import sys
import time
import argparse
import logging
from datetime import datetime, timezone
from typing import Dict, List, Any, Optional, Tuple
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock
from dotenv import load_dotenv
import jsonschema

# Load environment variables from root .env file
load_dotenv()

# Configure logging
logger = logging.getLogger(__name__)

# Constants following RULES.md standards
MAX_RETRIES = 3
DEFAULT_POLLING_INTERVAL = 300
DEFAULT_BATCH_SIZE = 100
DEFAULT_MAX_WORKERS = 5
DEFAULT_REQUESTS_PER_SECOND = 3.0
DEFAULT_BURST_SIZE = 10
DEFAULT_OPERATION_BATCH_SIZE = 10
DEFAULT_BACKOFF_SECONDS = 2.0

# Configuration schema for validation
CONFIG_SCHEMA = {
    "type": "object",
    "required": ["notion", "sync", "threading", "field_mapping"],
    "properties": {
        "notion": {
            "type": "object",
            "required": ["api_token", "source_database", "target_database"],
            "properties": {
                "api_token": {"type": "string"},
                "api_base": {"type": "string"},
                "version": {"type": "string"},
                "source_database": {"type": "string"},
                "target_database": {"type": "string"}
            }
        },
        "sync": {
            "type": "object",
            "properties": {
                "polling_interval_seconds": {"type": "integer", "minimum": 1},
                "batch_size": {"type": "integer", "minimum": 1},
                "max_retries": {"type": "integer", "minimum": 1},
                "backoff_seconds": {"type": "number", "minimum": 0}
            }
        },
        "threading": {
            "type": "object",
            "properties": {
                "max_workers": {"type": "integer", "minimum": 1},
                "requests_per_second": {"type": "number", "minimum": 0.1},
                "burst_size": {"type": "integer", "minimum": 1},
                "parallel_fetching": {"type": "boolean"},
                "parallel_operations": {"type": "boolean"},
                "operation_batch_size": {"type": "integer", "minimum": 1}
            }
        }
    }
}


class RateLimiter:
    """Thread-safe rate limiter using token bucket algorithm."""
    
    def __init__(self, requests_per_second: float, burst_size: Optional[int] = None) -> None:
        """Initialize rate limiter.
        
        Args:
            requests_per_second: Maximum sustained request rate
            burst_size: Maximum burst capacity (defaults to requests_per_second)
        """
        self.rate = requests_per_second
        self.burst_size = burst_size or int(requests_per_second)
        self.tokens = float(self.burst_size)
        self.lock = Lock()
        self.last_update = time.monotonic()
        
    def acquire(self, timeout: Optional[float] = None) -> bool:
        """Acquire permission to make a request.
        
        Args:
            timeout: Maximum time to wait for a token (None = wait forever)
            
        Returns:
            True if token acquired, False if timeout
        """
        deadline = time.monotonic() + timeout if timeout else float('inf')
        
        while time.monotonic() < deadline:
            with self.lock:
                now = time.monotonic()
                elapsed = now - self.last_update
                self.last_update = now
                
                # Add tokens based on elapsed time
                self.tokens = min(self.burst_size, self.tokens + elapsed * self.rate)
                
                if self.tokens >= 1:
                    self.tokens -= 1
                    return True
            
            # Calculate wait time
            wait_time = (1 - self.tokens) / self.rate
            time.sleep(min(wait_time, 0.01))
        
        return False


class NotionArticlesSync:
    """Sync articles from source database to archive database with threading support."""
    
    def __init__(self, config_path: str, debug: bool = False) -> None:
        """Initialize the sync with configuration.
        
        Args:
            config_path: Path to configuration JSON file
            debug: Enable debug logging mode
        """
        self.config = self.load_config(config_path)
        self.debug = debug
        
        # Setup logging
        self.setup_logging(debug, config_path)
        
        # Validate configuration
        self.validate_config()
        
        # Process environment variables in config
        self.process_environment_variables()
        
        self.headers = {
            "Authorization": f"Bearer {self.config['notion']['api_token']}",
            "Content-Type": "application/json",
            "Notion-Version": self.config['notion'].get('version', '2022-06-28')
        }
        
        self.source_db = self.config['notion']['source_database']
        self.target_db = self.config['notion']['target_database']
        self.api_base = self.config['notion'].get('api_base', 'https://api.notion.com/v1')
        
        # Sync configuration
        self.polling_interval = self.config['sync'].get('polling_interval_seconds', DEFAULT_POLLING_INTERVAL)
        self.batch_size = self.config['sync'].get('batch_size', DEFAULT_BATCH_SIZE)
        self.max_retries = self.config['sync'].get('max_retries', MAX_RETRIES)
        
        # Threading configuration
        threading_config = self.config.get('threading', {})
        self.max_workers = threading_config.get('max_workers', DEFAULT_MAX_WORKERS)
        self.requests_per_second = threading_config.get('requests_per_second', DEFAULT_REQUESTS_PER_SECOND)
        self.burst_size = threading_config.get('burst_size', DEFAULT_BURST_SIZE)
        self.parallel_fetching = threading_config.get('parallel_fetching', True)
        self.parallel_operations = threading_config.get('parallel_operations', True)
        self.operation_batch_size = threading_config.get('operation_batch_size', DEFAULT_OPERATION_BATCH_SIZE)
        
        # Initialize rate limiter
        self.rate_limiter = RateLimiter(self.requests_per_second, self.burst_size)
        
        # Field mappings
        self.field_mapping = self.config['field_mapping']['source_to_target']
        
        # Cache for database schemas
        self._schema_cache: Dict[str, Dict[str, Any]] = {}
        
        # Thread-safe API call counter
        self.api_call_count = 0
        self.api_call_lock = Lock()
        
        # Track last sync time
        self.last_sync_time: Optional[datetime] = None
        
        # Load last sync time from file if it exists
        self.sync_time_file = os.path.join(os.path.dirname(config_path), "notion_articles_sync_last_time.txt")
        self.load_sync_time()
        
        # Thread pool executor
        self.executor = ThreadPoolExecutor(max_workers=self.max_workers)
        
        logger.info("✅ Notion sync initialized")
        logger.info(f"📊 Source: {self.source_db}")
        logger.info(f"📊 Target: {self.target_db}")
        logger.info(f"🚀 Threading: {self.max_workers} workers, {self.requests_per_second} req/s")
        
        self.test_connections()
    
    def close(self) -> None:
        """Shut down the thread pool executor.

        Call explicitly (or use the object as a context manager) rather than
        relying on __del__ — during interpreter shutdown the logging system may
        already be torn down, making garbage-collection-time cleanup unreliable.
        """
        executor = getattr(self, 'executor', None)
        if executor is not None:
            executor.shutdown(wait=True)

    def __enter__(self) -> "NotionArticlesSync":
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        self.close()
    
    def setup_logging(self, debug: bool, config_path: str) -> None:
        """Setup logging configuration following RULES.md standards.
        
        Args:
            debug: Enable debug logging mode
            config_path: Path to configuration file for log file location
        """
        logger.setLevel(logging.DEBUG if debug else logging.INFO)
        
        # Console handler
        console_handler = logging.StreamHandler(sys.stdout)
        console_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        console_handler.setFormatter(console_formatter)
        logger.addHandler(console_handler)
        
        # File handler for debug mode
        if debug:
            # Get log file path (same directory as config, same name as script)
            config_dir = os.path.dirname(os.path.abspath(config_path))
            log_file = os.path.join(config_dir, "notion_articles_sync.log")
            
            file_handler = logging.FileHandler(log_file, encoding='utf-8')
            file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            file_handler.setFormatter(file_formatter)
            logger.addHandler(file_handler)
            
            logger.info(f"📝 Debug logging to file: {log_file}")
    
    def load_config(self, config_path: str) -> Dict[str, Any]:
        """Load configuration from JSON file.
        
        Args:
            config_path: Path to configuration file
            
        Returns:
            Configuration dictionary
            
        Raises:
            FileNotFoundError: If configuration file doesn't exist
            json.JSONDecodeError: If configuration file is invalid JSON
        """
        if not os.path.exists(config_path):
            logger.error(f"❌ Configuration file not found: {config_path}")
            sys.exit(1)
        
        try:
            with open(config_path, 'r') as f:
                return json.load(f)
        except json.JSONDecodeError as e:
            logger.error(f"❌ Invalid JSON in configuration file: {e}")
            sys.exit(1)
    
    def validate_config(self) -> None:
        """Validate configuration against schema.
        
        Raises:
            jsonschema.ValidationError: If configuration is invalid
        """
        try:
            jsonschema.validate(instance=self.config, schema=CONFIG_SCHEMA)
            logger.info("✅ Configuration validation passed")
        except jsonschema.ValidationError as e:
            logger.error(f"❌ Configuration validation failed: {e}")
            sys.exit(1)
    
    def process_environment_variables(self) -> None:
        """Process environment variables in configuration values.
        
        Currently only processes the NOTION_API_TOKEN environment variable.
        All other configuration values are taken directly from the JSON file.
        """
        def replace_env_vars(obj: Any) -> Any:
            """Recursively replace environment variables in configuration values."""
            if isinstance(obj, dict):
                return {k: replace_env_vars(v) for k, v in obj.items()}
            elif isinstance(obj, list):
                return [replace_env_vars(item) for item in obj]
            elif isinstance(obj, str) and obj.startswith('${') and obj.endswith('}'):
                env_var = obj[2:-1]
                value = os.getenv(env_var)
                if value is None:
                    logger.error(f"❌ Environment variable not found: {env_var}")
                    sys.exit(1)
                return value
            return obj
        
        self.config = replace_env_vars(self.config)
        logger.info("✅ Environment variables processed")
    
    def test_connections(self) -> None:
        """Test access to both databases."""
        logger.info("🔍 Testing database connections...")
        
        # Test source
        if self.api_call(f"databases/{self.source_db}"):
            logger.info("✅ Source database accessible")
        else:
            logger.error("❌ Cannot access source database")
            sys.exit(1)
        
        # Test target
        if self.api_call(f"databases/{self.target_db}"):
            logger.info("✅ Target database accessible")
        else:
            logger.error("❌ Cannot access target database")
            sys.exit(1)
    
    def increment_api_counter(self) -> int:
        """Thread-safe increment of API call counter."""
        with self.api_call_lock:
            self.api_call_count += 1
            return self.api_call_count
    
    def api_call(self, endpoint: str, method: str = "GET", data: Optional[dict] = None, 
                 retry: int = 0) -> Optional[dict]:
        """Make a Notion API call with retry logic and rate limiting."""
        url = f"{self.api_base}/{endpoint}"
        
        # Acquire rate limit token
        if not self.rate_limiter.acquire(timeout=30):
            logger.warning("⏱️ Rate limiter timeout")
            return None
        
        # Increment API call counter
        call_number = self.increment_api_counter()
        
        try:
            logger.debug(f"🌐 API Call #{call_number}: {method} {url}")
            if data:
                logger.debug(f"📦 Request data: {json.dumps(data, indent=2)}")
            
            if method == "GET":
                response = requests.get(url, headers=self.headers, timeout=30)
            elif method == "POST":
                response = requests.post(url, headers=self.headers, json=data, timeout=30)
            elif method == "PATCH":
                response = requests.patch(url, headers=self.headers, json=data, timeout=30)
            elif method == "DELETE":
                response = requests.delete(url, headers=self.headers, timeout=30)
            else:
                raise ValueError(f"Unsupported method: {method}")
            
            response.raise_for_status()
            return response.json()
            
        except requests.exceptions.RequestException as e:
            if hasattr(e, 'response') and e.response is not None:
                logger.debug(f"❌ Response status: {e.response.status_code}")
                try:
                    error_details = e.response.json()
                    logger.debug(f"❌ Error details: {json.dumps(error_details, indent=2)}")
                except ValueError:
                    logger.debug(f"❌ Response text: {e.response.text}")
            
            if retry < self.max_retries:
                wait = 2 ** retry
                logger.warning(f"⏳ Retrying in {wait}s (attempt {retry + 1}/{self.max_retries})")
                time.sleep(wait)
                return self.api_call(endpoint, method, data, retry + 1)
            else:
                logger.error(f"❌ API call failed: {e}")
                return None
    
    def normalize_rowid(self, rowid: str) -> str:
        """Remove dashes from rowid for consistent format.
        
        Args:
            rowid: Row ID to normalize
            
        Returns:
            Normalized row ID without dashes
        """
        if rowid:
            return rowid.replace("-", "")
        return rowid
    
    def extract_value(self, prop: Dict[str, Any]) -> Any:
        """Extract value from a Notion property.
        
        Args:
            prop: Notion property dictionary
            
        Returns:
            Extracted value from the property
        """
        prop_type = prop.get("type")
        
        if prop_type == "title":
            return "".join([t.get("plain_text", "") for t in prop.get("title", [])])
        elif prop_type == "rich_text":
            return "".join([t.get("plain_text", "") for t in prop.get("rich_text", [])])
        elif prop_type == "select":
            select = prop.get("select")
            return select.get("name") if select else None
        elif prop_type == "multi_select":
            return [opt.get("name") for opt in prop.get("multi_select", [])]
        elif prop_type == "date":
            date_obj = prop.get("date")
            return date_obj.get("start") if date_obj else None
        elif prop_type == "checkbox":
            return prop.get("checkbox", False)
        elif prop_type == "url":
            return prop.get("url")
        elif prop_type == "number":
            return prop.get("number")
        elif prop_type == "created_time":
            return prop.get("created_time")
        elif prop_type == "last_edited_time":
            return prop.get("last_edited_time")
        elif prop_type == "formula":
            formula = prop.get("formula", {})
            formula_type = formula.get("type")
            if formula_type == "string":
                return formula.get("string")
            elif formula_type == "number":
                return formula.get("number")
            elif formula_type == "boolean":
                return formula.get("boolean")
            elif formula_type == "date":
                date_obj = formula.get("date")
                return date_obj.get("start") if date_obj else None
        elif prop_type == "relation":
            return [rel.get("id") for rel in prop.get("relation", [])]
        
        return None
    
    def format_property(self, field_name: str, value: Any) -> Dict[str, Any]:
        """Format a property value for Notion API.
        
        Args:
            field_name: Name of the field to format
            value: Value to format
            
        Returns:
            Formatted property dictionary for Notion API
        """
        # Get target schema
        if self.target_db not in self._schema_cache:
            result = self.api_call(f"databases/{self.target_db}")
            if result:
                self._schema_cache[self.target_db] = result.get("properties", {})
        
        schema = self._schema_cache.get(self.target_db, {})
        field_type = schema.get(field_name, {}).get("type", "rich_text")
        
        # Format based on type
        if field_type == "title":
            return {"title": [{"text": {"content": str(value)}}]}
        elif field_type == "rich_text":
            return {"rich_text": [{"text": {"content": str(value)}}]}
        elif field_type == "select":
            return {"select": {"name": str(value)}}
        elif field_type == "multi_select":
            if isinstance(value, list):
                return {"multi_select": [{"name": str(v)} for v in value]}
            return {"multi_select": [{"name": str(value)}]}
        elif field_type == "date":
            if isinstance(value, str):
                return {"date": {"start": value}}
            return {"date": {"start": str(value)}}
        elif field_type == "checkbox":
            return {"checkbox": bool(value)}
        elif field_type == "number":
            try:
                return {"number": float(value)}
            except (ValueError, TypeError):
                return {"number": 0}
        elif field_type == "url":
            return {"url": str(value)}
        else:
            return {"rich_text": [{"text": {"content": str(value)}}]}
    
    def map_properties(self, source_props: Dict[str, Any]) -> Dict[str, Any]:
        """Map properties from source to target format.
        
        Args:
            source_props: Source properties dictionary
            
        Returns:
            Mapped properties dictionary for target format
        """
        mapped = {}
        
        for source_field, target_field in self.field_mapping.items():
            if source_field in source_props:
                value = self.extract_value(source_props[source_field])
                
                # Normalize rowid values
                if source_field == "rowid" and value:
                    value = self.normalize_rowid(value)
                
                if value is not None:
                    mapped[target_field] = self.format_property(target_field, value)
        
        return mapped
    
    def query_database(self, database_id: str, start_cursor: Optional[str] = None, 
                      filter_after: Optional[datetime] = None) -> Optional[Dict[str, Any]]:
        """Query a Notion database with optional time filtering.
        
        Args:
            database_id: ID of the database to query
            start_cursor: Optional cursor for pagination
            filter_after: Optional datetime filter for changes after this time
            
        Returns:
            Database query result or None if failed
        """
        data = {"page_size": self.batch_size}
        
        if start_cursor:
            data["start_cursor"] = start_cursor
        elif filter_after:
            data["filter"] = {
                "timestamp": "last_edited_time",
                "last_edited_time": {
                    "after": filter_after.isoformat()
                }
            }
            logger.debug(f"🔍 Filtering for changes after {filter_after.isoformat()}")
        
        return self.api_call(f"databases/{database_id}/query", method="POST", data=data)
    
    def fetch_all_items(self, database_id: str, filter_after: Optional[datetime] = None,
                       show_progress: bool = False) -> List[Dict[str, Any]]:
        """Fetch all items from a database with optional time filtering (sequential).
        
        Args:
            database_id: ID of the database to fetch from
            filter_after: Optional datetime filter for changes after this time
            show_progress: Whether to show progress information
            
        Returns:
            List of all items from the database
        """
        items = []
        has_more = True
        start_cursor = None
        page_num = 0
        
        while has_more:
            page_num += 1
            if show_progress:
                logger.info(f"📥 Fetching page {page_num} (API calls so far: {self.api_call_count})...")
            
            result = self.query_database(database_id, start_cursor, filter_after)
            if not result:
                break
            
            page_items = result.get("results", [])
            items.extend(page_items)
            
            if show_progress:
                logger.info(f"📊 Got {len(page_items)} items (total: {len(items)})")
            
            has_more = result.get("has_more", False)
            start_cursor = result.get("next_cursor")
        
        return items
    
    def get_last_sync_time(self) -> Optional[datetime]:
        """Get the last sync time from the tracked sync timestamp."""
        if self.last_sync_time:
            logger.info(f"📅 Last sync time: {self.last_sync_time.isoformat()}")
            return self.last_sync_time
        
        logger.info("📅 No previous sync found - will perform full sync")
        return None
    
    def set_last_sync_time(self, sync_time: Optional[datetime] = None) -> None:
        """Set the last sync time manually.
        
        Args:
            sync_time: The sync time to set. If None, sets to current UTC time.
        """
        if sync_time is None:
            sync_time = datetime.now(timezone.utc)
        
        self.last_sync_time = sync_time
        self.save_sync_time()
        logger.info(f"📅 Manually set last sync time to: {self.last_sync_time.isoformat()}")
    
    def reset_sync_time(self) -> None:
        """Reset the last sync time to force a full sync on next run."""
        self.last_sync_time = None
        self.save_sync_time()
        logger.info("📅 Reset last sync time - next sync will be full sync")
    
    def load_sync_time(self) -> None:
        """Load the last sync time from file."""
        try:
            if os.path.exists(self.sync_time_file):
                with open(self.sync_time_file, 'r') as f:
                    timestamp_str = f.read().strip()
                    if timestamp_str:
                        self.last_sync_time = datetime.fromisoformat(timestamp_str)
                        logger.debug(f"📅 Loaded last sync time: {self.last_sync_time.isoformat()}")
        except Exception as e:
            logger.warning(f"⚠️ Could not load sync time from file: {e}")
            self.last_sync_time = None
    
    def save_sync_time(self) -> None:
        """Save the current sync time to file."""
        try:
            if self.last_sync_time:
                with open(self.sync_time_file, 'w') as f:
                    f.write(self.last_sync_time.isoformat())
                logger.debug(f"📅 Saved sync time to file: {self.last_sync_time.isoformat()}")
        except Exception as e:
            logger.warning(f"⚠️ Could not save sync time to file: {e}")
    
    def show_sync_status(self) -> None:
        """Show the current sync status."""
        if self.last_sync_time:
            logger.info(f"📅 Current sync status:")
            logger.info(f"   └─ Last sync: {self.last_sync_time.isoformat()}")
            logger.info(f"   └─ Sync file: {self.sync_time_file}")
        else:
            logger.info("📅 Current sync status: No previous sync found")
    
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
    
    def create_archive_entry(self, source_item: Dict[str, Any]) -> Optional[str]:
        """Create a new entry in the archive database.
        
        Args:
            source_item: Source Notion item to archive
            
        Returns:
            Archive item ID if successful, None otherwise
        """
        source_rowid = self.extract_value(source_item.get("properties", {}).get("rowid", {}))
        logger.debug(f"📝 Creating archive entry for: {source_rowid}")
        
        mapped_props = self.map_properties(source_item.get("properties", {}))
        
        data = {
            "parent": {"database_id": self.target_db},
            "properties": mapped_props
        }
        
        result = self.api_call("pages", method="POST", data=data)
        if result:
            archive_id = result.get("id")
            return archive_id
        
        return None
    
    def update_archive_entry(self, archive_item: Dict[str, Any], source_item: Dict[str, Any]) -> bool:
        """Update existing archive entry.
        
        Args:
            archive_item: Target archive item to update
            source_item: Source item with new data
            
        Returns:
            True if update successful, False otherwise
        """
        logger.debug(f"🔄 Updating archive entry: {archive_item.get('id')}")
        
        mapped_props = self.map_properties(source_item.get("properties", {}))
        data = {"properties": mapped_props}
        
        result = self.api_call(f"pages/{archive_item['id']}", method="PATCH", data=data)
        return result is not None
    
    def delete_archive_entry(self, archive_item: Dict[str, Any]) -> bool:
        """Delete or archive an entry when source is deleted.
        
        Args:
            archive_item: Archive item to delete/archive
            
        Returns:
            True if deletion successful, False otherwise
        """
        logger.debug(f"🗑️  Archiving deleted entry: {archive_item.get('id')}")
        
        data = {"archived": True}
        
        result = self.api_call(f"pages/{archive_item['id']}", method="PATCH", data=data)
        return result is not None
    
    def update_source_tracking(self, source_id: str, archive_id: str,
                               current_target_rowid: Optional[str] = None) -> bool:
        """Update source database with archive tracking.

        Skips the PATCH (and the resulting source.last_edited_time bump that fuels
        the incremental-sync feedback loop) when the source already tracks this
        archive id.

        Args:
            source_id: Source item ID to update
            archive_id: Archive item ID to track
            current_target_rowid: Existing 'target rowid' value on the source item, if known.
                Pass to short-circuit no-op writes.

        Returns:
            True if no write was needed or the write succeeded, False on API failure.
        """
        normalized_id = self.normalize_rowid(archive_id)

        if current_target_rowid and self.normalize_rowid(current_target_rowid) == normalized_id:
            return True

        data = {
            "properties": {
                "target rowid": {"rich_text": [{"text": {"content": normalized_id}}]}
            }
        }

        result = self.api_call(f"pages/{source_id}", method="PATCH", data=data)
        return result is not None
    
    def process_batch_operations(self, operations: List[Tuple[Any, ...]], 
                                operation_type: str) -> Tuple[int, int]:
        """Process a batch of operations in parallel.
        
        Args:
            operations: List of operations to perform
            operation_type: Type of operation ('create', 'update', 'delete')
            
        Returns:
            Tuple of (successful_count, failed_count)
        """
        if not operations:
            return 0, 0
        
        success_count = 0
        failed_count = 0
        
        if not self.parallel_operations or len(operations) == 1:
            # Process sequentially
            for op in operations:
                try:
                    if operation_type == 'create':
                        item = op[0]
                        archive_id = self.create_archive_entry(item)
                        if archive_id:
                            self.update_source_tracking(item["id"], archive_id)
                            success_count += 1
                        else:
                            failed_count += 1
                    elif operation_type == 'update':
                        source_item, target_item = op
                        if self.update_archive_entry(target_item, source_item):
                            current_tracking = self.extract_value(
                                source_item.get("properties", {}).get("target rowid", {})
                            )
                            self.update_source_tracking(
                                source_item["id"], target_item["id"], current_tracking
                            )
                            success_count += 1
                        else:
                            failed_count += 1
                    elif operation_type == 'delete':
                        item = op[0]
                        if self.delete_archive_entry(item):
                            success_count += 1
                        else:
                            failed_count += 1
                except Exception as e:
                    logger.error(f"❌ Operation failed: {e}")
                    failed_count += 1
        else:
            # Process in parallel batches
            with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                futures = []
                
                for op in operations:
                    if operation_type == 'create':
                        future = executor.submit(self._process_create_operation, op[0])
                    elif operation_type == 'update':
                        future = executor.submit(self._process_update_operation, op[0], op[1])
                    elif operation_type == 'delete':
                        future = executor.submit(self._process_delete_operation, op[0])
                    
                    futures.append(future)
                
                # Wait for all operations to complete
                for future in as_completed(futures):
                    try:
                        result = future.result(timeout=60)
                        if result:
                            success_count += 1
                        else:
                            failed_count += 1
                    except Exception as e:
                        logger.error(f"❌ Batch operation failed: {e}")
                        failed_count += 1
        
        return success_count, failed_count
    
    def _process_create_operation(self, item: Dict[str, Any]) -> bool:
        """Process a single create operation.
        
        Args:
            item: Source item to create archive entry for
            
        Returns:
            True if operation successful, False otherwise
        """
        try:
            archive_id = self.create_archive_entry(item)
            if archive_id:
                self.update_source_tracking(item["id"], archive_id)
                return True
        except Exception as e:
            logger.error(f"❌ Create failed: {e}")
        return False
    
    def _process_update_operation(self, source_item: Dict[str, Any], target_item: Dict[str, Any]) -> bool:
        """Process a single update operation.

        Args:
            source_item: Source item with new data
            target_item: Target archive item to update

        Returns:
            True if operation successful, False otherwise
        """
        try:
            if self.update_archive_entry(target_item, source_item):
                current_tracking = self.extract_value(
                    source_item.get("properties", {}).get("target rowid", {})
                )
                self.update_source_tracking(
                    source_item["id"], target_item["id"], current_tracking
                )
                return True
        except Exception as e:
            logger.error(f"❌ Update failed: {e}")
        return False
    
    def _process_delete_operation(self, item: Dict[str, Any]) -> bool:
        """Process a single delete operation.
        
        Args:
            item: Archive item to delete/archive
            
        Returns:
            True if operation successful, False otherwise
        """
        try:
            return self.delete_archive_entry(item)
        except Exception as e:
            logger.error(f"❌ Delete failed: {e}")
        return False
    
    def run_sync(self, force_full_sync: bool = False) -> bool:
        """Run a single sync cycle."""
        logger.info("=" * 80)
        logger.info("🚀 Starting sync cycle" + (" (FULL SYNC)" if force_full_sync else " (INCREMENTAL)"))
        
        try:
            # Reset API counter
            with self.api_call_lock:
                self.api_call_count = 0
            
            # Detect changes
            if force_full_sync:
                new_items, updated_items, deleted_items = self.detect_changes_full_sync()
            else:
                new_items, updated_items, deleted_items = self.detect_changes_incremental()
            
            # Process items if there are changes
            total_changes = len(new_items) + len(updated_items) + len(deleted_items)
            
            if total_changes == 0:
                logger.info("✅ No changes detected - databases are in sync")
                return True
            
            logger.info(f"📋 Processing {total_changes} changes...")
            
            # Process operations in batches
            total_success = 0
            total_failed = 0
            
            # Process new items in batches
            if new_items:
                logger.info(f"📝 Creating {len(new_items)} new items...")
                for i in range(0, len(new_items), self.operation_batch_size):
                    batch = new_items[i:i + self.operation_batch_size]
                    batch_ops = [(item,) for item in batch]
                    success, failed = self.process_batch_operations(batch_ops, 'create')
                    total_success += success
                    total_failed += failed
                    logger.info(f"   └─ Batch {i//self.operation_batch_size + 1}: "
                              f"{success} succeeded, {failed} failed")
            
            # Process updated items in batches
            if updated_items:
                logger.info(f"🔄 Updating {len(updated_items)} items...")
                for i in range(0, len(updated_items), self.operation_batch_size):
                    batch = updated_items[i:i + self.operation_batch_size]
                    success, failed = self.process_batch_operations(batch, 'update')
                    total_success += success
                    total_failed += failed
                    logger.info(f"   └─ Batch {i//self.operation_batch_size + 1}: "
                              f"{success} succeeded, {failed} failed")
            
            # Process deleted items in batches
            if deleted_items:
                logger.info(f"🗑️  Deleting {len(deleted_items)} items...")
                for i in range(0, len(deleted_items), self.operation_batch_size):
                    batch = deleted_items[i:i + self.operation_batch_size]
                    batch_ops = [(item,) for item in batch]
                    success, failed = self.process_batch_operations(batch_ops, 'delete')
                    total_success += success
                    total_failed += failed
                    logger.info(f"   └─ Batch {i//self.operation_batch_size + 1}: "
                              f"{success} succeeded, {failed} failed")
            
            logger.info(f"✅ Sync completed: {total_success} succeeded, {total_failed} failed")
            logger.info(f"📊 Total API calls: {self.api_call_count}")
            
            # Update last sync time on successful completion
            if total_failed == 0:
                self.last_sync_time = datetime.now(timezone.utc)
                self.save_sync_time()
                logger.info(f"📅 Updated last sync time to: {self.last_sync_time.isoformat()}")
            
            return total_failed == 0
            
        except Exception as e:
            logger.error(f"❌ Sync cycle failed: {e}")
            if self.debug:
                import traceback
                logger.debug(traceback.format_exc())
            return False
    
    def run_continuous(self, force_full_sync_first: bool = False) -> None:
        """Run continuous sync with polling.
        
        Args:
            force_full_sync_first: Whether to force a full sync on the first run
        """
        logger.info(f"🔄 Starting continuous sync (every {self.polling_interval}s)")
        
        first_run = True
        while True:
            try:
                # Full sync on first run if requested, then incremental
                self.run_sync(force_full_sync=force_full_sync_first and first_run)
                first_run = False
                
                logger.info(f"💤 Waiting {self.polling_interval} seconds...")
                time.sleep(self.polling_interval)
                
            except KeyboardInterrupt:
                logger.info("⏹️  Sync stopped by user")
                break
            except Exception as e:
                logger.error(f"❌ Error: {e}")
                logger.info("💤 Waiting 60s before retry...")
                time.sleep(60)


def main():
    """Main function."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    default_config = os.path.join(script_dir, "notion_articles_sync.json")
    
    parser = argparse.ArgumentParser(description="Sync Notion articles to archive")
    parser.add_argument("--config", type=str, default=default_config,
                        help="Path to configuration JSON file")
    parser.add_argument("--once", action="store_true", default=False,
                        help="Run sync once and exit")
    parser.add_argument("--full-sync", action="store_true", default=False,
                        help="Force a full sync (check for deletions)")
    parser.add_argument("--reset-sync-time", action="store_true", default=False,
                        help="Reset the last sync time to force a full sync")
    parser.add_argument("--status", action="store_true", default=False,
                        help="Show current sync status and exit")
    parser.add_argument("--debug", action="store_true", default=False,
                        help="Enable debug logging and log to file")
    
    args = parser.parse_args()
    
    try:
        # Initialize and run sync (context manager guarantees the thread pool
        # executor is shut down explicitly, not at garbage-collection time)
        with NotionArticlesSync(args.config, debug=args.debug) as sync:
            if args.reset_sync_time:
                sync.reset_sync_time()

            if args.status:
                sync.show_sync_status()
                sys.exit(0)

            if args.once:
                sync.run_sync(force_full_sync=args.full_sync)
            else:
                sync.run_continuous(force_full_sync_first=args.full_sync)

    except KeyboardInterrupt:
        logger.info("⏹️ Sync stopped")
    except Exception as e:
        logger.error(f"❌ Fatal error: {e}")
        if args.debug:
            import traceback
            logger.debug(traceback.format_exc())
        sys.exit(1)


if __name__ == "__main__":
    main()