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
import sys
from typing import Dict, List, Optional, Tuple
from urllib.parse import urlparse, urlunparse

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
        return notion_utils.query_recently_created_pages(self.headers, self.database_id, days)

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
        return notion_utils.patch_page_properties(self.headers, page_id, {"link": {"url": new_url}})

    def process_database(self, days: int, dry_run: bool = False, testing_mode: bool = False) -> List[Dict]:
        """
        Process the database and update URLs.
        
        Args:
            days: Number of days to look back.
            dry_run: If True, no changes are written to Notion.
            testing_mode: If True, validates that cleaned URLs are accessible.
        """
        def validation_note(cleaned_url: str) -> str:
            is_valid, status_msg = self._check_url_validity(cleaned_url)
            return f" [{'✅' if is_valid else '❌'} Link: {status_msg}]"

        return notion_utils.normalize_pages(
            self._query_notion_database(days),
            extract_page_info=self._extract_page_info,
            transform=self._clean_url,
            update_page=self._update_page_url,
            make_result=lambda page_id, _last_edited, original, cleaned: {
                'page_id': page_id,
                'original_url': original,
                'cleaned_url': cleaned,
            },
            dry_run=dry_run,
            verb="clean",
            unchanged_label="Already clean",
            annotate=validation_note if testing_mode else None,
        )


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

