#!/usr/bin/env python3
"""
Notion Article Name Normalizer

Automatically normalizes article names in a Notion database to sentence case
while preserving proper names, acronyms, and special tokens.

Note: ALL CAPS words with 2+ characters are preserved as-is (assumed to be
acronyms, emphasis, or intentionally capitalized terms). This means "TEST ALL CAPS"
will remain "TEST ALL CAPS" rather than becoming "Test all caps".
"""

import argparse
import json
import logging
import os
import re
import sys
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple, Any

import requests
from dotenv import load_dotenv

# Load environment variables from root .env file
load_dotenv()


class NotionNameNormalizer:
    """Main class for normalizing Notion article names."""
    
    def __init__(self, config_path: str):
        """Initialize with configuration file path."""
        self.config = self._load_config(config_path)
        
        # Get config values
        self.notion_api_key = os.getenv('NOTION_API_TOKEN') or self.config.get('notion_api_key')
        self.database_id = self.config.get('database_id')
        
        # Validate required config values
        if not all([self.notion_api_key, self.database_id]):
            raise ValueError("Missing required configuration values")
        
        # Set up Notion API headers
        self.headers = {
            'Authorization': f'Bearer {self.notion_api_key}',
            'Notion-Version': '2022-06-28',
            'Content-Type': 'application/json'
        }
        
        # Optional spaCy support
        self.spacy_nlp = None
        if self.config.get('use_spacy', False):
            try:
                import spacy
                self.spacy_nlp = spacy.load("en_core_web_sm")
                logging.info("🧠 spaCy loaded successfully for entity detection")
            except ImportError:
                logging.warning("⚠️  spaCy requested but not available, falling back to heuristics")
        
        logging.info("✅ Name normalizer initialized")
        logging.info(f"📊 Database ID: {self.database_id}")
    
    def _load_config(self, config_path: str) -> Dict:
        """Load and parse the JSON configuration file."""
        try:
            with open(config_path, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            # Try looking in the same directory as this script
            script_dir = os.path.dirname(os.path.abspath(__file__))
            fallback_path = os.path.join(script_dir, os.path.basename(config_path))
            try:
                with open(fallback_path, 'r') as f:
                    logging.info(f"📁 Loaded config from fallback path: {fallback_path}")
                    return json.load(f)
            except FileNotFoundError:
                raise FileNotFoundError(f"Configuration file not found at {config_path} or {fallback_path}")
            except json.JSONDecodeError:
                raise ValueError(f"Invalid JSON in fallback configuration file: {fallback_path}")
        except json.JSONDecodeError:
            raise ValueError(f"Invalid JSON in configuration file: {config_path}")
    
    def _query_notion_database(self, days: int) -> List[Dict]:
        """Query Notion database for articles created in the last N days."""
        pages = []
        
        # Calculate the date filter
        filter_date = datetime.utcnow() - timedelta(days=days)
        filter_date_str = filter_date.isoformat() + 'Z'
        
        logging.info(f"🔍 Querying database {self.database_id} for articles created in the last {days} days (since {filter_date_str})")
        
        # Build the filter for created_time and sort by created_time descending
        filter_body = {
            "filter": {
                "and": [
                    {
                        "property": "created",
                        "created_time": {
                            "after": filter_date_str
                        }
                    }
                ]
            },
            "sorts": [
                {
                    "property": "created",
                    "direction": "descending"
                }
            ]
        }
        
        start_cursor = None
        page_count = 0
        
        while True:
            # Add pagination cursor if we have one
            if start_cursor:
                filter_body["start_cursor"] = start_cursor
            
            try:
                response = requests.post(
                    f"https://api.notion.com/v1/databases/{self.database_id}/query",
                    headers=self.headers,
                    json=filter_body
                )
                response.raise_for_status()
                
                data = response.json()
                results = data.get('results', [])
                
                if not results:
                    break
                
                pages.extend(results)
                page_count += len(results)
                logging.debug(f"📥 Retrieved {len(results)} pages (total: {page_count})")
                
                # Check if there are more pages
                if not data.get('has_more', False):
                    break
                
                start_cursor = data.get('next_cursor')
                
            except requests.exceptions.RequestException as e:
                logging.error(f"❌ Notion API error: {e}")
                if hasattr(e, 'response') and e.response is not None:
                    logging.error(f"📊 Status code: {e.response.status_code}")
                    logging.error(f"📄 Response: {e.response.text[:200]}...")
                raise
        
        logging.info(f"📊 Total pages retrieved: {len(pages)}")
        return pages
    
    def _extract_page_info(self, page: Dict) -> Optional[Tuple[str, str, str]]:
        """Extract page ID, last edited time, and name from a Notion page."""
        page_id = page.get('id', '')
        last_edited_time = page.get('last_edited_time', '')
        
        # Extract name from article title property
        properties = page.get('properties', {})
        name_property = properties.get('article', {})
        
        if not name_property or name_property.get('type') != 'title':
            logging.warning(f"⚠️  Page {page_id} missing or invalid article property")
            return None
        
        title_content = name_property.get('title', [])
        if not title_content:
            logging.warning(f"⚠️  Page {page_id} has empty article property")
            return None
        
        # Extract plain text from rich text
        name_text = ''.join([segment.get('plain_text', '') for segment in title_content])
        if not name_text.strip():
            logging.warning(f"⚠️  Page {page_id} has empty article text")
            return None
        
        return page_id, last_edited_time, name_text.strip()
    
    def _normalize_name(self, original_name: str) -> str:
        """Normalize the name to sentence case while preserving proper names and special tokens."""
        if not original_name:
            return original_name
        
        # Store original tokens for comparison during restoration
        original_tokens = original_name.split()
        
        # Step 1: Apply sentence case - capitalize first character only
        normalized = self._apply_sentence_case(original_name)
        
        # Step 2: Process tokens to restore proper names from whitelist and handle sentence boundaries
        normalized_tokens = normalized.split()
        result_tokens = []
        i = 0
        
        while i < len(normalized_tokens):
            # Check if we can form a multi-word proper name starting at position i
            multi_word_found = False
            whitelist = self.config.get('proper_name_whitelist', [])
            
            for proper_name in whitelist:
                if ' ' in proper_name:  # Only multi-word names
                    proper_words = proper_name.lower().split()
                    proper_length = len(proper_words)
                    
                    # Check if we have enough tokens remaining and they match
                    if (i + proper_length <= len(normalized_tokens) and 
                        [word.lower() for word in normalized_tokens[i:i + proper_length]] == proper_words):
                        # Found a multi-word proper name, add it and skip ahead
                        result_tokens.extend(proper_name.split())
                        i += proper_length
                        multi_word_found = True
                        logging.debug(f"🔄 Restored multi-word proper name: {' '.join(normalized_tokens[i:i+proper_length])} -> {proper_name}")
                        break
            
            if not multi_word_found:
                # Process as single token - check whitelist and apply preservation rules
                orig_token = original_tokens[i]
                norm_token = normalized_tokens[i]
                
                # Check if this token should be capitalized due to sentence boundaries
                # Only capitalize if it's the first token or follows sentence-ending punctuation
                should_capitalize = self._should_capitalize_token(i, normalized_tokens, original_tokens)
                
                if should_capitalize and norm_token and norm_token[0].isalpha():
                    norm_token = norm_token[0].upper() + norm_token[1:]
                else:
                    # Ensure non-sentence-starting tokens are lowercase
                    if norm_token and norm_token[0].isalpha():
                        norm_token = norm_token[0].lower() + norm_token[1:]
                
                restored_token = self._restore_token_capitalization(orig_token, norm_token)
                result_tokens.append(restored_token)
                i += 1
        
        # Step 3: Reconstruct string and normalize whitespace
        result = ' '.join(result_tokens)
        result = re.sub(r'\s+', ' ', result).strip()
        
        return result
    
    def _should_capitalize_token(self, token_index: int, normalized_tokens: List[str], original_tokens: List[str]) -> bool:
        """Determine if a token should be capitalized based on sentence boundaries."""
        if token_index == 0:
            return True  # Always capitalize first token
        
        # Check if previous token ends with sentence-ending punctuation
        prev_token = normalized_tokens[token_index - 1]
        if re.search(r'[.!?]$', prev_token):
            return True
        
        return False
    
    def _apply_sentence_case(self, text: str) -> str:
        """Apply sentence case: capitalize first letter and after sentence boundaries."""
        if not text:
            return text
        
        # Only capitalize the first letter of the entire text
        # Words after proper names should remain lowercase
        if text and text[0].isalpha():
            text = text[0].upper() + text[1:]
        
        return text
    

    
    def _restore_token_capitalization(self, original_token: str, normalized_token: str) -> str:
        """Restore proper capitalization for a specific token based on preservation rules."""
        # Rule 1: Preserve ALL CAPS words (2+ chars) - assumed to be acronyms/emphasis
        # This prevents "TEST ALL CAPS" from becoming "Test all caps"
        if len(original_token) >= 2 and original_token.isupper():
            return original_token
        
        # Rule 2: Preserve tokens with special punctuation (e.g., "U.S.A.", "A&B")
        if re.search(r'[.&-]', original_token):
            return original_token
        
        # Rule 3: Check against user-defined proper name whitelist from config
        whitelist = self.config.get('proper_name_whitelist', [])
        for proper_name in whitelist:
            if self._is_token_match(original_token, proper_name):
                return proper_name
        
        # Rule 4: Use spaCy NLP to detect person names if available
        if self.spacy_nlp:
            if self._is_person_entity(original_token):
                return original_token
        
        # If no preservation rules apply, use the normalized (sentence case) version
        return normalized_token
    
    def _is_token_match(self, token: str, proper_name: str) -> bool:
        """Check if a token matches a proper name (case-insensitive word boundary match)."""
        # Convert both to lowercase for case-insensitive comparison
        token_lower = token.lower()
        proper_lower = proper_name.lower()
        
        # Check for exact match first (most common case)
        if token_lower == proper_lower:
            return True
        
        # Check if token is a component of a multi-word proper name
        # e.g., "John" matches "John Smith" in the whitelist
        if token_lower in proper_lower.split():
            return True
        
        return False
    
    def _is_person_entity(self, token: str) -> bool:
        """Check if a token is recognized as a person entity by spaCy NLP."""
        # Early return if spaCy is not available
        if not self.spacy_nlp:
            return False
        
        # Process token through spaCy NLP pipeline
        doc = self.spacy_nlp(token)
        
        # Check if any detected entity is labeled as a person
        for ent in doc.ents:
            if ent.label_ == 'PERSON':
                return True
        
        return False
    
    def _update_page_property(self, page_id: str, property_name: str, new_value: str) -> bool:
        """Update a specific property of a Notion page."""
        try:
            # Prepare the update payload for the title property
            update_data = {
                "properties": {
                    property_name: {
                        "title": [
                            {
                                "type": "text",
                                "text": {
                                    "content": new_value
                                }
                            }
                        ]
                    }
                }
            }
            
            response = requests.patch(
                f"https://api.notion.com/v1/pages/{page_id}",
                headers=self.headers,
                json=update_data
            )
            response.raise_for_status()
            
            logging.debug(f"✅ Successfully updated page {page_id[:8]}...{page_id[-2:]}")
            return True
            
        except requests.exceptions.RequestException as e:
            logging.error(f"❌ Failed to update page {page_id[:8]}...{page_id[-2:]}: {e}")
            if hasattr(e, 'response') and e.response is not None:
                logging.error(f"📊 Status code: {e.response.status_code}")
                logging.error(f"📄 Response: {e.response.text[:200]}...")
            return False
        except Exception as e:
            logging.error(f"❌ Unexpected error updating page {page_id[:8]}...{page_id[-2:]}: {e}")
            return False
    
    def process_database(self, days: int) -> List[Dict]:
        """Main method to process the database and update normalized results."""
        # Query Notion database for pages created in the specified time period
        pages = self._query_notion_database(days)
        
        # Initialize counters and results collection
        results = []
        processed_count = 0
        updated_count = 0
        unchanged_count = 0
        
        for page in pages:
            # Extract essential page information (ID, timestamp, article name)
            page_info = self._extract_page_info(page)
            if not page_info:
                continue
            
            page_id, last_edited_time, original_name = page_info
            # Apply normalization rules to the article name
            normalized_name = self._normalize_name(original_name)
            
            # Store results for reporting
            result = {
                'page_id': page_id,
                'last_edited_time': last_edited_time,
                'original_name': original_name,
                'normalized_name': normalized_name
            }
            
            results.append(result)
            processed_count += 1
            
            # Determine if normalization actually changed the name
            if original_name != normalized_name:
                # Attempt to update the Notion page with normalized name
                if self._update_page_property(page_id, 'article', normalized_name):
                    updated_count += 1
                    logging.info(f'📝 Article: "{original_name}" normalized to "{normalized_name}"')
                else:
                    logging.error(f'❌ Article: "{original_name}" failed to update to "{normalized_name}"')
            else:
                # Name was already in expected format (e.g., ALL CAPS preserved)
                unchanged_count += 1
                logging.info(f'✅ Article: "{original_name}" did not change, was already normalized')
            
            # Log first few normalizations at DEBUG level for troubleshooting
            if processed_count <= 3:
                logging.debug(f"🔍 Sample normalization: '{original_name}' -> '{normalized_name}'")
        
        # Log final processing summary
        logging.info(f"✅ Successfully processed {processed_count} pages")
        logging.info(f"📝 Updated {updated_count} pages with normalized names")
        logging.info(f"✅ Left unchanged {unchanged_count} pages (already normalized)")
        return results
    

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
        description="Normalize Notion article names to sentence case while preserving proper names"
    )
    parser.add_argument(
        '--days', 
        type=int, 
        default=1,
        help='Number of days to look back for created articles (default: 1)'
    )
    parser.add_argument(
        '--config', 
        type=str,
        default="normalize_names.json",
        help='Path to JSON config file (default: "normalize_names.json")'
    )
    parser.add_argument(
        '--debug', 
        action='store_true',
        help='Enable debug logging'
    )
    
    args = parser.parse_args()
    
    # Set up logging
    setup_logging(args.debug)
    
    try:
        # Initialize normalizer
        normalizer = NotionNameNormalizer(args.config)
        
        # Process database and update properties
        results = normalizer.process_database(args.days)
        
        
    except Exception as e:
        logging.error(f"❌ Fatal error: {e}")
        if args.debug:
            logging.exception("Full traceback:")
        sys.exit(1)


if __name__ == '__main__':
    main()