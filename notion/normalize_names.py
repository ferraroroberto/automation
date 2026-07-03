#!/usr/bin/env python3
"""
Notion Article Name Normalizer

Automatically normalizes article names in a Notion database to sentence case
while preserving proper names, acronyms, and special tokens.

Usage:
    python normalize_names.py --days 14 --config normalize_names.json
    python normalize_names.py --days 7 --dry-run  # Preview changes without updating
    python normalize_names.py --test "TEST ALL CAPS"  # Test mode
"""

import argparse
import json
import logging
import os
import re
import sys
from datetime import datetime, timedelta, timezone
from typing import Dict, List, Optional, Tuple, Any

import requests
from dotenv import load_dotenv

# Load environment variables from root .env file
load_dotenv()


class NotionNameNormalizer:
    """
    Main class for normalizing Notion article names.
    
    Preserves ALL CAPS words (2+ chars) as acronyms/emphasis.
    Example: "TEST ALL CAPS" remains "TEST ALL CAPS"
    """
    
    def __init__(self, config_path: str):
        """
        Initialize normalizer with configuration file path.
        
        Args:
            config_path: Path to JSON configuration file
        """
        self.config = self._load_config(config_path)
        self._setup_api_credentials()
        self._load_word_lists()
        self._initialize_spacy()
        self._log_initialization()
    
    def _setup_api_credentials(self):
        """Setup API credentials and headers."""
        self.notion_api_key = os.getenv('NOTION_API_TOKEN') or self.config.get('notion_api_key')
        self.database_id = self.config.get('database_id')
        
        if not all([self.notion_api_key, self.database_id]):
            raise ValueError("Missing required configuration values: notion_api_key and database_id")
        
        self.headers = {
            'Authorization': f'Bearer {self.notion_api_key}',
            'Notion-Version': '2022-06-28',
            'Content-Type': 'application/json'
        }
    
    def _load_word_lists(self):
        """Load word lists from configuration."""
        self.proper_names = set(self.config.get('proper_name_whitelist', []))
        self.special_cases = set(self.config.get('special_cases', []))
        self.common_words = set(self.config.get('common_words', []))
        self.common_words_with_punct = set(self.config.get('common_words_with_punct', []))
    
    def _initialize_spacy(self):
        """Initialize optional spaCy NLP support."""
        self.spacy_nlp = None
        if self.config.get('use_spacy', False):
            try:
                import spacy
                logging.info("🔄 Loading spaCy model (this may take a moment)...")
                self.spacy_nlp = spacy.load("en_core_web_sm")
                logging.info("🧠 spaCy loaded successfully for entity detection")
            except ImportError:
                logging.warning("⚠️  spaCy requested but not available, falling back to heuristics")
            except Exception as e:
                logging.warning(f"⚠️  spaCy failed to load: {e}. Falling back to heuristics")
                self.spacy_nlp = None
    
    def _log_initialization(self):
        """Log initialization summary."""
        logging.info("✅ Name normalizer initialized")
        logging.info(f"📊 Database ID: {self.database_id}")
        logging.info(f"🔤 Special cases: {len(self.special_cases)}, Common words: {len(self.common_words)}")
    
    def _load_config(self, config_path: str) -> Dict:
        """Load and parse the JSON configuration file."""
        paths_to_try = [
            config_path,
            os.path.join(os.path.dirname(os.path.abspath(__file__)), os.path.basename(config_path))
        ]
        
        for path in paths_to_try:
            try:
                with open(path, 'r') as f:
                    config = json.load(f)
                if path != config_path:
                    logging.info(f"📁 Loaded config from fallback path: {path}")
                return config
            except FileNotFoundError:
                continue
            except json.JSONDecodeError:
                raise ValueError(f"Invalid JSON in configuration file: {path}")
        
        raise FileNotFoundError(f"Configuration file not found at {' or '.join(paths_to_try)}")
    
    def _query_notion_database(self, days: int) -> List[Dict]:
        """Query Notion database for articles created in the last N days."""
        filter_date = datetime.now(timezone.utc).replace(tzinfo=None) - timedelta(days=days)
        filter_date_str = filter_date.isoformat() + 'Z'
        
        logging.info(f"🔍 Querying database for articles from last {days} days (since {filter_date_str})")
        
        filter_body = {
            "filter": {"and": [{"property": "created", "created_time": {"after": filter_date_str}}]},
            "sorts": [{"property": "created", "direction": "descending"}]
        }
        
        pages = []
        start_cursor = None
        
        while True:
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
                logging.debug(f"📥 Retrieved {len(results)} pages (total: {len(pages)})")
                
                if not data.get('has_more', False):
                    break
                
                start_cursor = data.get('next_cursor')
                
            except requests.exceptions.RequestException as e:
                logging.error(f"❌ Notion API error: {e}")
                if hasattr(e, 'response') and e.response:
                    logging.error(f"📊 Status: {e.response.status_code}, Response: {e.response.text[:200]}...")
                raise
        
        logging.info(f"📊 Total pages retrieved: {len(pages)}")
        return pages
    
    def _extract_page_info(self, page: Dict) -> Optional[Tuple[str, str, str]]:
        """Extract essential information from a Notion page object."""
        page_id = page.get('id', '')
        last_edited_time = page.get('last_edited_time', '')
        
        # Extract article name from the 'article' property (title type)
        properties = page.get('properties', {})
        name_property = properties.get('article', {})
        
        if not name_property or name_property.get('type') != 'title':
            logging.warning(f"⚠️  Page {page_id} missing or invalid article property")
            return None
        
        title_content = name_property.get('title', [])
        if not title_content:
            logging.warning(f"⚠️  Page {page_id} has empty article property")
            return None
        
        name_text = ''.join([segment.get('plain_text', '') for segment in title_content]).strip()
        if not name_text:
            logging.warning(f"⚠️  Page {page_id} has empty article text")
            return None
        
        return page_id, last_edited_time, name_text
    
    def _normalize_name(self, original_name: str) -> str:
        """
        Normalize article name to sentence case while preserving proper names and special tokens.
        
        Args:
            original_name: Original article name to normalize
            
        Returns:
            Normalized name in sentence case
        """
        if not original_name:
            return original_name
        
        # Apply basic sentence case and process tokens
        normalized = self._apply_sentence_case(original_name)
        original_tokens = original_name.split()
        normalized_tokens = normalized.split()
        
        result_tokens = []
        i = 0
        
        while i < len(normalized_tokens):
            # Try to match multi-word proper names first
            matched_length = self._find_multi_word_match(normalized_tokens, i)
            
            if matched_length > 1:
                # Found multi-word proper name - add it and skip ahead
                proper_name = self._get_proper_name_for_tokens(normalized_tokens[i:i + matched_length])
                result_tokens.extend(proper_name.split())
                i += matched_length
            else:
                # Process single token
                orig_token = original_tokens[i]
                norm_token = normalized_tokens[i]
                
                # Extract sentence-ending punctuation from current token
                sentence_punct, word_part = self._extract_sentence_punctuation(norm_token)
                
                # Apply capitalization rules (check previous token for sentence-ending punctuation)
                should_cap = self._should_capitalize_token(i, normalized_tokens)
                
                # Normalize the word part
                if should_cap:
                    word_part = self._capitalize_first_letter(word_part)
                else:
                    word_part = self._lowercase_first_letter(word_part)
                
                # Restore proper capitalization for special cases (using original token)
                restored_word = self._restore_token_capitalization(orig_token, word_part)
                
                # Reattach sentence-ending punctuation
                final_token = restored_word + sentence_punct
                result_tokens.append(final_token)
                i += 1
        
        return re.sub(r'\s+', ' ', ' '.join(result_tokens)).strip()
    
    def _find_multi_word_match(self, tokens: List[str], start_index: int) -> int:
        """Find the length of a multi-word proper name match starting at the given index."""
        for proper_name in self.proper_names:
            if ' ' not in proper_name:  # Skip single-word names
                continue
                
            proper_words = proper_name.lower().split()
            proper_length = len(proper_words)
            
            if (start_index + proper_length <= len(tokens) and 
                [word.lower() for word in tokens[start_index:start_index + proper_length]] == proper_words):
                if self._is_valid_multi_word_match(tokens[start_index:start_index + proper_length], proper_name):
                    return proper_length
        return 1  # No multi-word match found
    
    def _get_proper_name_for_tokens(self, tokens: List[str]) -> str:
        """Get the proper capitalization for a sequence of tokens."""
        token_text = ' '.join([t.lower() for t in tokens])
        for proper_name in self.proper_names:
            if proper_name.lower() == token_text:
                return proper_name
        return ' '.join(tokens)  # Fallback
    
    def _should_capitalize_token(self, token_index: int, normalized_tokens: List[str]) -> bool:
        """
        Determine if a token should be capitalized based on sentence boundaries.
        
        Checks if previous token ends with sentence-ending punctuation.
        Acronyms like U.S.A. should not trigger capitalization.
        """
        if token_index == 0:
            return True
        
        prev_token = normalized_tokens[token_index - 1]
        
        # Check if previous token is an acronym (has internal periods)
        # If so, don't treat its final period as sentence-ending
        if re.search(r'[A-Za-z]\.[A-Za-z]', prev_token):
            acronym_match = re.match(r'^[A-Za-z](?:\.[A-Za-z])+\.?$', prev_token)
            if acronym_match:
                # It's an acronym - only capitalize if it has ! or ? after it
                return bool(re.search(r'[!?]+$', prev_token))
        
        # Regular sentence-ending punctuation
        return bool(re.search(r'[.!?]+$', prev_token))
    
    def _capitalize_first_letter(self, token: str) -> str:
        """Capitalize the first letter of a token if it's alphabetic."""
        return token[0].upper() + token[1:] if token and token[0].isalpha() else token
    
    def _lowercase_first_letter(self, token: str) -> str:
        """Lowercase the first letter of a token if it's alphabetic."""
        return token[0].lower() + token[1:] if token and token[0].isalpha() else token
    
    def _apply_sentence_case(self, text: str) -> str:
        """Apply basic sentence case: capitalize first letter only."""
        if not text or not text[0].isalpha():
            return text
        return text[0].upper() + text[1:].lower()
    
    def _restore_token_capitalization(self, original_token: str, normalized_token: str) -> str:
        """
        Restore proper capitalization for a specific token based on preservation rules.
        
        Note: This works on the word part without sentence-ending punctuation.
        """
        # Extract sentence-ending punctuation from original token for comparison
        _, orig_word_part = self._extract_sentence_punctuation(original_token)
        
        # Rule 1: Preserve ALL CAPS words (2+ chars) - check alphabetic part only
        alpha_part = re.sub(r'[^\w]', '', orig_word_part)
        if len(alpha_part) >= 2 and alpha_part.isupper():
            return orig_word_part
        
        # Rule 2: Preserve the pronoun "I"
        if orig_word_part.lower() == 'i':
            return 'I'
        
        # Rule 3: Check against proper name whitelist
        if self._is_proper_name(orig_word_part):
            return self._get_proper_name_capitalization(orig_word_part)
        
        # Rule 4: Use spaCy NLP to detect person names if available
        if self.spacy_nlp and self._is_person_entity(orig_word_part):
            return orig_word_part
        
        # Rule 5: Handle tokens with other punctuation (not sentence-ending)
        if re.search(r'[^\w\s]', orig_word_part):
            return self._handle_punctuated_token(orig_word_part, normalized_token)
        
        return normalized_token
    
    def _is_proper_name(self, token: str) -> bool:
        """Check if token is a known proper name."""
        token_lower = token.lower()
        # Check for exact match or as part of multi-word proper name
        for proper_name in self.proper_names:
            if proper_name.lower() == token_lower:
                return True
            if ' ' in proper_name and token_lower in proper_name.lower().split() and token_lower not in self.common_words:
                return True
        return False
    
    def _get_proper_name_capitalization(self, token: str) -> str:
        """Get the correct capitalization for a proper name token."""
        token_lower = token.lower()
        for proper_name in self.proper_names:
            if proper_name.lower() == token_lower:
                return proper_name
            if ' ' in proper_name:
                proper_words = proper_name.split()
                for word in proper_words:
                    if word.lower() == token_lower:
                        return word
        return token
    
    def _handle_punctuated_token(self, token: str, normalized_token: str) -> str:
        """
        Handle tokens that contain punctuation (excluding sentence-ending punctuation).
        
        This handles contractions, apostrophes, commas, etc.
        """
        # Extract first alphabetic part (before any punctuation)
        alpha_part_match = re.match(r'^([A-Za-z]+)', token)
        if not alpha_part_match:
            return token
        
        alpha_part = alpha_part_match.group(1)
        
        # Preserve ALL CAPS words with punctuation (e.g., "U.S.A.")
        if len(alpha_part) >= 2 and alpha_part.isupper():
            return token
        
        # Preserve proper names with punctuation
        if self._is_proper_name(alpha_part):
            return token
        
        # Extract normalized alphabetic part (normalized_token might have punctuation too)
        normalized_alpha_match = re.match(r'^([A-Za-z]+)', normalized_token)
        normalized_alpha = normalized_alpha_match.group(1) if normalized_alpha_match else normalized_token
        
        # Reconstruct with normalized word part
        return self._reconstruct_with_punctuation(normalized_alpha, token)

    def _reconstruct_with_punctuation(self, new_alpha_part: str, original_token: str) -> str:
        """
        Reconstruct a token with punctuation, replacing the alphabetic part.
        
        For contractions like "Here's", we need to preserve the structure.
        Example: reconstruct_with_punctuation("here", "Here's") -> "here's"
        """
        # For contractions (apostrophes), preserve the original structure
        # Example: "Here's" -> normalize "Here" to "here", keep "'s"
        if "'" in original_token:
            # Split on apostrophe
            parts = original_token.split("'", 1)
            if len(parts) == 2:
                first_part = parts[0]
                rest = "'" + parts[1]
                # Replace first part with normalized version
                return new_alpha_part + rest
        
        # For other punctuation, use regex to find and replace first alphabetic sequence
        # Match first sequence of letters
        match = re.match(r'^([A-Za-z]+)', original_token)
        if match:
            return new_alpha_part + original_token[len(match.group(1)):]
        
        # Fallback: try to replace first alphabetic part
        parts = re.findall(r'[^\w]+|\w+', original_token)
        result_parts = []
        alpha_replaced = False
        
        for part in parts:
            if part.isalpha() and not alpha_replaced:
                result_parts.append(new_alpha_part)
                alpha_replaced = True
            else:
                result_parts.append(part)
        
        return ''.join(result_parts)
    
    def _extract_sentence_punctuation(self, token: str) -> Tuple[str, str]:
        """
        Extract sentence-ending punctuation from a token.
        
        Only extracts if punctuation is truly sentence-ending (not part of acronym like U.S.A.).
        Acronyms typically have periods between letters, not just at the end.
        
        Returns:
            Tuple of (punctuation_suffix, word_part)
            Example: ("!", "Amazing") from "Amazing!"
        """
        # Check if token looks like an acronym (has periods between letters, e.g., U.S.A.)
        # Works with both uppercase (U.S.A.) and lowercase (u.s.a.) after normalization
        # Pattern: letter.letter.letter (with optional final period)
        if re.search(r'[A-Za-z]\.[A-Za-z]', token):
            # Check if there's sentence punctuation AFTER the acronym pattern
            # Match pattern like "U.S.A." or "U.S.A.!" or "u.s.a." 
            acronym_match = re.match(r'^([A-Za-z](?:\.[A-Za-z])+\.?)([.!?]*)$', token)
            if acronym_match:
                # It's an acronym - don't extract the periods as sentence punctuation
                # But extract any trailing ! or ? after the acronym
                base = acronym_match.group(1)
                trailing_punct = acronym_match.group(2)
                if trailing_punct and re.search(r'[!?]', trailing_punct):
                    # Has ! or ? after acronym - extract those
                    return trailing_punct, base
                return "", token
        
        # Extract trailing sentence-ending punctuation
        match = re.search(r'([.!?]+)$', token)
        if match:
            punct = match.group(1)
            word_part = token[:-len(punct)]
            return punct, word_part
        return "", token

    def _is_valid_multi_word_match(self, tokens: List[str], proper_name: str) -> bool:
        """
        Validate if a multi-word match is actually a proper name and not just common words.
        """
        token_text = ' '.join([t.lower() for t in tokens])
        proper_lower = proper_name.lower()
        
        # Exact match is always valid
        if token_text == proper_lower:
            return True
        
        # Count distinctive (non-common) words in the match
        distinctive_words = [word for word in tokens if word.lower() not in self.common_words]
        
        # Require at least 2 distinctive words or be in special cases
        return len(distinctive_words) >= 2 or proper_name in self.special_cases
    
    def _is_person_entity(self, token: str) -> bool:
        """Check if a token is recognized as a person entity by spaCy NLP."""
        if not self.spacy_nlp:
            return False
        
        doc = self.spacy_nlp(token)
        return any(ent.label_ == 'PERSON' for ent in doc.ents)
    
    def _update_page_property(self, page_id: str, property_name: str, new_value: str) -> bool:
        """Update a specific property of a Notion page via API."""
        update_data = {
            "properties": {
                property_name: {
                    "title": [{"type": "text", "text": {"content": new_value}}]
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
            logging.debug(f"✅ Successfully updated page {page_id[:8]}...{page_id[-2:]}")
            return True
            
        except requests.exceptions.RequestException as e:
            page_short = f"{page_id[:8]}...{page_id[-2:]}"
            logging.error(f"❌ Failed to update page {page_short}: {e}")
            if hasattr(e, 'response') and e.response:
                logging.error(f"📊 Status: {e.response.status_code}, Response: {e.response.text[:200]}...")
            return False
        except Exception as e:
            logging.error(f"❌ Unexpected error updating page {page_id[:8]}...{page_id[-2:]}: {e}")
            return False
    
    def process_database(self, days: int, dry_run: bool = False) -> List[Dict]:
        """
        Main method to process the database and update normalized results.
        
        Args:
            days: Number of days to look back for articles
            dry_run: If True, only show what would be changed without updating Notion
            
        Returns:
            List of processing results with original and normalized names
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
            
            page_id, last_edited_time, original_name = page_info
            normalized_name = self._normalize_name(original_name)
            
            result = {
                'page_id': page_id,
                'last_edited_time': last_edited_time,
                'original_name': original_name,
                'normalized_name': normalized_name
            }
            results.append(result)
            stats['processed'] += 1
            
            # Update if changed, otherwise log as unchanged
            if original_name != normalized_name:
                if dry_run:
                    stats['would_update'] += 1
                    logging.info(f'📝 [DRY RUN] Would change: "{original_name}" → "{normalized_name}"')
                else:
                    if self._update_page_property(page_id, 'article', normalized_name):
                        stats['updated'] += 1
                        logging.info(f'📝 "{original_name}" → "{normalized_name}"')
                    else:
                        logging.error(f'❌ Failed to update: "{original_name}"')
            else:
                stats['unchanged'] += 1
                logging.info(f'✅ Already normalized: "{original_name}"')
            
            # Log sample normalizations for debugging
            if stats['processed'] <= 3:
                logging.debug(f"🔍 Sample: '{original_name}' → '{normalized_name}'")
        
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
    """
    Set up logging configuration with appropriate level and format.
    
    Args:
        debug: If True, enable DEBUG level logging
    """
    level = logging.DEBUG if debug else logging.INFO
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[logging.StreamHandler(sys.stdout)]
    )


def main():
    """
    Main entry point with command line argument parsing.
    
    Supports:
    - --days: Number of days to look back (default: 14)
    - --config: Path to config file (default: normalize_names.json)
    - --debug: Enable debug logging
    - --test: Test mode with specific string
    - --dry-run: Preview changes without updating Notion
    """
    parser = argparse.ArgumentParser(
        description="Normalize Notion article names to sentence case while preserving proper names"
    )
    parser.add_argument(
        '--days', 
        type=int, 
        default=14,
        help='Number of days to look back for created articles (default: 14)'
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
    parser.add_argument(
        '--test',
        type=str,
        help='Test mode: normalize a specific string without processing the database'
    )
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Dry run mode: show what would be changed without updating Notion'
    )
    
    args = parser.parse_args()
    
    try:
        # Set up logging first (user must explicitly use --debug if they want detailed output)
        setup_logging(args.debug)
        
        # Test mode: normalize a specific string without database processing
        if args.test:
            # Initialize normalizer for testing
            normalizer = NotionNameNormalizer(args.config)
            
            logging.info("🧪 Test mode: normalizing string: '%s'", args.test)
            normalized = normalizer._normalize_name(args.test)
            logging.info("Original:   %s", args.test)
            logging.info("Normalized: %s", normalized)
            logging.info("Changed:    %s", 'Yes' if args.test != normalized else 'No')
            return
        
        # Initialize normalizer for database processing
        normalizer = NotionNameNormalizer(args.config)
        
        # Process database and update article properties (or dry run)
        results = normalizer.process_database(args.days, dry_run=args.dry_run)
        
        # Log successful completion
        logging.info("✅ Script completed successfully")
        
    except Exception as e:
        logging.error(f"❌ Fatal error: {e}")
        if args.debug:
            logging.exception("Full traceback:")
        sys.exit(1)


if __name__ == '__main__':
    main()