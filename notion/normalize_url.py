#!/usr/bin/env python3
"""
Notion URL Normalizer

Automatically cleans URLs in a Notion database by removing unnecessary query parameters
(like UTM tags) while preserving parameters for specific domains (e.g., YouTube).

Usage:
    python normalize_url.py --days 14 --config normalize_url.json
    python normalize_url.py --days 7 --dry-run
    python normalize_url.py --test "https://example.com?utm_source=test"
"""

import argparse
import logging
import os
import sys
from datetime import datetime, timedelta, timezone
from typing import Dict, List, Optional, Tuple
from urllib.parse import urlparse, urlunparse, parse_qs, urlencode

import requests
from dotenv import load_dotenv

import utils as notion_utils

# Load environment variables
load_dotenv()


class NotionURLNormalizer:
    """
    Main class for normalizing Notion article URLs.
    """
    
    def __init__(self, config_path: str):
        """Initialize normalizer with configuration file path."""
        self.config = self._load_config(config_path)
        self._setup_api_credentials()
        self.domains_preserving_params = set(self.config.get('domains_preserving_params', []))
        self._log_initialization()
    
    def _setup_api_credentials(self):
        """Setup API credentials and headers."""
        self.notion_api_key, self.database_id, self.headers = notion_utils.resolve_notion_credentials(self.config)
    
    def _log_initialization(self):
        """Log initialization summary."""
        logging.info("✅ URL normalizer initialized")
        logging.info(f"📊 Database ID: {self.database_id}")
        logging.info(f"🛡️  Preserved domains: {len(self.domains_preserving_params)}")

    def _load_config(self, config_path: str) -> Dict:
        """Load and parse the JSON configuration file."""
        return notion_utils.load_json_config_with_fallback(config_path, __file__)

    def _clean_url(self, original_url: str) -> str:
        """
        Remove query parameters from URL unless domain is in preservation list.
        
        Args:
            original_url: The URL to clean.
            
        Returns:
            Cleaned URL.
        """
        if not original_url:
            return original_url
            
        try:
            parsed = urlparse(original_url)
            
            # Check if domain is in the preserved list
            domain = parsed.netloc.lower()
            # Handle subdomains roughly
            if any(p in domain for p in self.domains_preserving_params):
                return original_url
                
            # Reconstruct URL without query parameters and fragment
            # scheme://netloc/path (drop params, query, fragment)
            # We keep 'params' (path parameters) usually, but query is what we want to strip.
            # urlunparse expects: scheme, netloc, path, params, query, fragment
            clean = urlunparse((parsed.scheme, parsed.netloc, parsed.path, parsed.params, '', ''))
            
            return clean
            
        except Exception as e:
            logging.warning(f"⚠️  Failed to parse URL '{original_url}': {e}")
            return original_url

    def _check_url_validity(self, url: str) -> Tuple[bool, str]:
        """
        Check if a URL is accessible using a HEAD (or fallback GET) request.
        
        Args:
            url: The URL to check.
            
        Returns:
            Tuple of (is_valid, status_message)
        """
        try:
            # Use a standard browser User-Agent to avoid some 403 blocks
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            }
            
            # Try HEAD request first (faster)
            response = requests.head(url, headers=headers, allow_redirects=True, timeout=10)
            
            # If Method Not Allowed (405), try GET (stream=True to download only headers/start)
            if response.status_code == 405:
                response = requests.get(url, headers=headers, stream=True, timeout=10)
            
            if 200 <= response.status_code < 400:
                return True, f"OK ({response.status_code})"
            else:
                return False, f"Error: {response.status_code}"
                
        except requests.RequestException as e:
            return False, f"Failed: {type(e).__name__}"

    def _query_notion_database(self, days: int) -> List[Dict]:
        """Query Notion database for articles created in the last N days."""
        filter_date = datetime.now(timezone.utc).replace(tzinfo=None) - timedelta(days=days)
        filter_date_str = filter_date.isoformat() + 'Z'

        logging.info(f"🔍 Querying database for articles from last {days} days (since {filter_date_str})")

        query_body = {
            "filter": {"and": [{"property": "created", "created_time": {"after": filter_date_str}}]},
            "sorts": [{"property": "created", "direction": "descending"}]
        }

        return notion_utils.paginated_database_query(self.headers, self.database_id, query_body)

    def _extract_page_info(self, page: Dict) -> Optional[Tuple[str, str, str]]:
        """Extract essential information (ID, time, URL) from a Notion page object."""
        page_id = page.get('id', '')
        last_edited_time = page.get('last_edited_time', '')
        
        properties = page.get('properties', {})
        
        # Check for 'link' property as defined in build_newsletter.py
        url_prop = properties.get('link', {})
        
        if not url_prop or url_prop.get('type') != 'url':
            # Try fallback or just log warning if strictly required
            # normalize_names checks strictly, so we will too but simpler
            return None
            
        url_content = url_prop.get('url', '')
        
        if not url_content:
            return None
            
        return page_id, last_edited_time, url_content

    def _update_page_url(self, page_id: str, new_url: str) -> bool:
        """Update the link property of a Notion page."""
        update_data = {
            "properties": {
                "link": {
                    "url": new_url
                }
            }
        }
        
        try:
            response = requests.patch(
                f"https://api.notion.com/v1/pages/{page_id}",
                headers=self.headers,
                json=update_data
            )
            response.raise_for_status()
            logging.debug(f"✅ Successfully updated page {page_id[:8]}...")
            return True
            
        except requests.exceptions.RequestException as e:
            logging.error(f"❌ Failed to update page {page_id[:8]}...: {e}")
            return False

    def process_database(self, days: int, dry_run: bool = False, testing_mode: bool = False) -> List[Dict]:
        """
        Process the database and update URLs.
        
        Args:
            days: Number of days to look back.
            dry_run: If True, no changes are written to Notion.
            testing_mode: If True, validates that cleaned URLs are accessible.
        """
        pages = self._query_notion_database(days)
        results = []
        stats = {'processed': 0, 'updated': 0, 'unchanged': 0, 'would_update': 0}
        
        if dry_run:
            logging.info("🔍 DRY RUN MODE: No changes will be made to Notion")
        
        for page in pages:
            page_info = self._extract_page_info(page)
            if not page_info:
                continue
            
            page_id, last_edited_time, original_url = page_info
            cleaned_url = self._clean_url(original_url)
            
            # Validate URL if testing mode is enabled
            validation_msg = ""
            if testing_mode:
                is_valid, status_msg = self._check_url_validity(cleaned_url)
                status_icon = "✅" if is_valid else "❌"
                validation_msg = f" [{status_icon} Link: {status_msg}]"
            
            result = {
                'page_id': page_id,
                'original_url': original_url,
                'cleaned_url': cleaned_url
            }
            results.append(result)
            stats['processed'] += 1
            
            if original_url != cleaned_url:
                if dry_run:
                    stats['would_update'] += 1
                    logging.info(f'📝 [DRY RUN] Would clean: "{original_url}" → "{cleaned_url}"{validation_msg}')
                else:
                    if self._update_page_url(page_id, cleaned_url):
                        stats['updated'] += 1
                        logging.info(f'📝 Cleaned: "{original_url}" → "{cleaned_url}"{validation_msg}')
                    else:
                        logging.error(f'❌ Failed to update: "{original_url}"')
            else:
                stats['unchanged'] += 1
                logging.info(f'✅ Already clean: "{original_url}"{validation_msg}')
        
        self._log_summary(stats, dry_run)
        return results

    def _log_summary(self, stats: Dict[str, int], dry_run: bool = False):
        """Log processing summary."""
        logging.info(f"✅ Processed {stats['processed']} pages")
        if dry_run:
            logging.info(f"📝 Would update {stats['would_update']}, unchanged {stats['unchanged']}")
        else:
            logging.info(f"📝 Updated {stats['updated']}, unchanged {stats['unchanged']}")


def setup_logging(debug: bool = False):
    """Set up logging configuration."""
    level = logging.DEBUG if debug else logging.INFO
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[logging.StreamHandler(sys.stdout)]
    )


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Normalize Notion article URLs by removing query parameters"
    )
    parser.add_argument(
        '--days', 
        type=int, 
        default=14,
        help='Number of days to look back (default: 14)'
    )
    parser.add_argument(
        '--config', 
        type=str, 
        default="normalize_url.json",
        help='Path to JSON config file'
    )
    parser.add_argument(
        '--debug', 
        action='store_true',
        help='Enable debug logging'
    )
    parser.add_argument(
        '--test',
        type=str,
        help='Test mode: clean a specific URL string'
    )
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Dry run mode'
    )
    parser.add_argument(
        '--testing',
        action='store_true',
        help='Test validity of cleaned URLs (checks if page exists)'
    )
    
    args = parser.parse_args()
    
    setup_logging(args.debug)
    
    try:
        if args.test:
            normalizer = NotionURLNormalizer(args.config)
            logging.info("🧪 Test mode: cleaning URL: '%s'", args.test)
            cleaned = normalizer._clean_url(args.test)
            logging.info("Original: %s", args.test)
            logging.info("Cleaned:  %s", cleaned)
            logging.info("Changed:  %s", 'Yes' if args.test != cleaned else 'No')

            if args.testing:
                is_valid, status_msg = normalizer._check_url_validity(cleaned)
                status_icon = "✅" if is_valid else "❌"
                logging.info("Validation: %s %s", status_icon, status_msg)
                
            return

        normalizer = NotionURLNormalizer(args.config)
        normalizer.process_database(args.days, dry_run=args.dry_run, testing_mode=args.testing)
        
    except Exception as e:
        logging.error(f"❌ Fatal error: {e}")
        if args.debug:
            logging.exception("Full traceback:")
        sys.exit(1)


if __name__ == '__main__':
    main()

