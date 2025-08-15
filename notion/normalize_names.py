#!/usr/bin/env python3
"""
Notion Article Name Normalizer

Automatically normalizes article names in a Notion database to sentence case
while preserving proper names, acronyms, and special tokens.
"""

import argparse
import json
import logging
import os
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any

import requests
from dotenv import load_dotenv

# Load environment variables from root .env file
load_dotenv()

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)


class NotionNameNormalizer:
    """Main class for normalizing Notion article names."""
    
    def __init__(self, config_path: str):
        """Initialize with configuration file path."""
        self.config = self._load_config(config_path)
        self.session = requests.Session()
        self.session.headers.update({
            'Authorization': f"Bearer {self.config['notion_api_key']}",
            'Notion-Version': '2022-06-28',
            'Content-Type': 'application/json'
        })
        
        # Optional spaCy support
        self.spacy_nlp = None
        if self.config.get('use_spacy', False):
            try:
                import spacy
                self.spacy_nlp = spacy.load("en_core_web_sm")
                logger.info("spaCy loaded successfully for entity detection")
            except ImportError:
                logger.warning("spaCy requested but not available, falling back to heuristics")
    
    def _load_config(self, config_path: str) -> Dict:
        """Load and validate configuration file."""
        try:
            with open(config_path, 'r') as f:
                config = json.load(f)
        except FileNotFoundError:
            raise FileNotFoundError(f"Configuration file not found: {config_path}")
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON in configuration file: {e}")
        
        # Process environment variables in config
        config = self._process_environment_variables(config)
        
        # Validate required keys
        required_keys = ['notion_api_key', 'database_id']
        missing_keys = [key for key in required_keys if key not in config]
        if missing_keys:
            raise ValueError(f"Missing required configuration keys: {missing_keys}")
        
        logger.info(f"Configuration loaded from: {config_path}")
        return config
    
    def _process_environment_variables(self, config: Dict) -> Dict:
        """Process environment variables in configuration values.
        
        Supports ${ENV_VAR_NAME} syntax for environment variable substitution.
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
                    logger.error(f"Environment variable not found: {env_var}")
                    sys.exit(1)
                return value
            return obj
        
        return replace_env_vars(config)
    
    def _query_notion_database(self, days: int) -> List[Dict]:
        """Query Notion database for pages edited in the last N days."""
        database_id = self.config['database_id']
        pages = []
        
        # Calculate the date filter
        filter_date = datetime.utcnow() - timedelta(days=days)
        filter_date_str = filter_date.isoformat() + 'Z'
        
        logger.info(f"Querying database {database_id} for pages edited since {filter_date_str}")
        
        # Build the filter for last_edited_time
        filter_body = {
            "filter": {
                "property": "last_edited_time",
                "last_edited_time": {
                    "past_n_days": days
                }
            }
        }
        
        start_cursor = None
        page_count = 0
        
        while True:
            # Add pagination cursor if we have one
            if start_cursor:
                filter_body["start_cursor"] = start_cursor
            
            try:
                response = self.session.post(
                    f"https://api.notion.com/v1/databases/{database_id}/query",
                    json=filter_body
                )
                response.raise_for_status()
                
                data = response.json()
                results = data.get('results', [])
                
                if not results:
                    break
                
                pages.extend(results)
                page_count += len(results)
                logger.debug(f"Retrieved {len(results)} pages (total: {page_count})")
                
                # Check if there are more pages
                if not data.get('has_more', False):
                    break
                
                start_cursor = data.get('next_cursor')
                
            except requests.exceptions.RequestException as e:
                logger.error(f"Notion API error: {e}")
                if hasattr(e, 'response') and e.response is not None:
                    logger.error(f"Status code: {e.response.status_code}")
                    logger.error(f"Response: {e.response.text[:200]}...")
                raise
        
        logger.info(f"Total pages retrieved: {len(pages)}")
        return pages
    
    def _extract_page_info(self, page: Dict) -> Optional[Tuple[str, str, str]]:
        """Extract page ID, last edited time, and name from a Notion page."""
        page_id = page.get('id', '')
        last_edited_time = page.get('last_edited_time', '')
        
        # Extract name from title property
        properties = page.get('properties', {})
        name_property = properties.get('Name', {})
        
        if not name_property or name_property.get('type') != 'title':
            logger.warning(f"Page {page_id} missing or invalid Name property")
            return None
        
        title_content = name_property.get('title', [])
        if not title_content:
            logger.warning(f"Page {page_id} has empty Name property")
            return None
        
        # Extract plain text from rich text
        name_text = ''.join([segment.get('plain_text', '') for segment in title_content])
        if not name_text.strip():
            logger.warning(f"Page {page_id} has empty Name text")
            return None
        
        return page_id, last_edited_time, name_text.strip()
    
    def _normalize_name(self, original_name: str) -> str:
        """Normalize the name to sentence case while preserving proper names and special tokens."""
        if not original_name:
            return original_name
        
        # Store original for comparison
        original_tokens = original_name.split()
        
        # Step 1: Convert to lowercase
        normalized = original_name.lower()
        
        # Step 2: Capitalize first character (sentence case)
        if normalized:
            normalized = normalized[0].upper() + normalized[1:]
        
        # Step 3: Restore proper names and special tokens
        normalized_tokens = normalized.split()
        
        for i, (orig_token, norm_token) in enumerate(zip(original_tokens, normalized_tokens)):
            restored_token = self._restore_token_capitalization(orig_token, norm_token)
            if restored_token != norm_token:
                normalized_tokens[i] = restored_token
                logger.debug(f"Restored token: '{norm_token}' -> '{restored_token}'")
        
        # Step 4: Reconstruct and clean up whitespace
        result = ' '.join(normalized_tokens)
        result = re.sub(r'\s+', ' ', result).strip()
        
        return result
    
    def _restore_token_capitalization(self, original_token: str, normalized_token: str) -> str:
        """Restore proper capitalization for a specific token."""
        # Check if token should remain ALL CAPS (length >= 2)
        if len(original_token) >= 2 and original_token.isupper():
            return original_token
        
        # Check if token contains special characters (periods, ampersands, hyphens)
        if re.search(r'[.&-]', original_token):
            return original_token
        
        # Check against proper name whitelist
        whitelist = self.config.get('proper_name_whitelist', [])
        for proper_name in whitelist:
            if self._is_token_match(original_token, proper_name):
                return proper_name
        
        # Use spaCy for entity detection if available
        if self.spacy_nlp:
            if self._is_person_entity(original_token):
                return original_token
        
        return normalized_token
    
    def _is_token_match(self, token: str, proper_name: str) -> bool:
        """Check if a token matches a proper name (case-insensitive word boundary match)."""
        # Simple word boundary matching
        token_lower = token.lower()
        proper_lower = proper_name.lower()
        
        # Exact match
        if token_lower == proper_lower:
            return True
        
        # Check if token is part of the proper name
        if token_lower in proper_lower.split():
            return True
        
        return False
    
    def _is_person_entity(self, token: str) -> bool:
        """Check if a token is recognized as a person entity by spaCy."""
        if not self.spacy_nlp:
            return False
        
        doc = self.spacy_nlp(token)
        for ent in doc.ents:
            if ent.label_ == 'PERSON':
                return True
        return False
    
    def process_database(self, days: int) -> List[Dict]:
        """Main method to process the database and return normalized results."""
        # Query Notion database
        pages = self._query_notion_database(days)
        
        results = []
        processed_count = 0
        
        for page in pages:
            page_info = self._extract_page_info(page)
            if not page_info:
                continue
            
            page_id, last_edited_time, original_name = page_info
            normalized_name = self._normalize_name(original_name)
            
            result = {
                'page_id': page_id,
                'last_edited_time': last_edited_time,
                'original_name': original_name,
                'normalized_name': normalized_name
            }
            
            results.append(result)
            processed_count += 1
            
            # Log at DEBUG level for sample items
            if processed_count <= 3:
                logger.debug(f"Sample normalization: '{original_name}' -> '{normalized_name}'")
        
        logger.info(f"Successfully processed {processed_count} pages")
        return results
    
    def save_results(self, results: List[Dict], output_path: str):
        """Save results to JSON output file."""
        try:
            with open(output_path, 'w') as f:
                json.dump(results, f, indent=2, ensure_ascii=False)
            logger.info(f"Results saved to: {output_path}")
        except Exception as e:
            logger.error(f"Failed to save results: {e}")
            raise
    
    def print_summary(self, results: List[Dict]):
        """Print summary to stdout."""
        print(f"\nMatched pages: {len(results)}")
        
        for result in results:
            page_id = result['page_id'][:8] + '...' + result['page_id'][-2:]  # Truncate for display
            last_edited = result['last_edited_time'][:19].replace('T', ' ')  # Format timestamp
            original = result['original_name']
            normalized = result['normalized_name']
            
            print(f"{page_id} | {last_edited} | {original} -> {normalized}")


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Normalize Notion article names to sentence case while preserving proper names"
    )
    parser.add_argument(
        '--days', 
        type=int, 
        default=7,
        help='Number of days to look back for edited pages (default: 7)'
    )
    parser.add_argument(
        '--config', 
        type=str,
        help='Path to configuration file (default: same basename as script)'
    )
    parser.add_argument(
        '--debug', 
        action='store_true',
        help='Enable debug logging'
    )
    
    args = parser.parse_args()
    
    # Set logging level
    if args.debug:
        logger.setLevel(logging.DEBUG)
        logger.debug("Debug logging enabled")
    
    # Determine config file path
    if args.config:
        config_path = args.config
    else:
        # Use same basename as script
        script_path = Path(__file__)
        config_path = script_path.with_suffix('.json')
    
    # Determine output file path
    script_path = Path(__file__)
    output_path = script_path.with_suffix('_output.json')
    
    try:
        # Initialize normalizer
        normalizer = NotionNameNormalizer(str(config_path))
        
        # Process database
        results = normalizer.process_database(args.days)
        
        # Save results
        normalizer.save_results(results, str(output_path))
        
        # Print summary
        normalizer.print_summary(results)
        
    except Exception as e:
        logger.error(f"Fatal error: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()