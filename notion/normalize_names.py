#!/usr/bin/env python3
"""
Notion Article Name Normalizer

Automatically normalizes article names in a Notion database to sentence case
while preserving proper names, acronyms, and special tokens.

Usage:
    python normalize_names.py --days 14 --config normalize_names.json
    python normalize_names.py --test "TEST ALL CAPS"  # Test mode
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
        
        # Extract required configuration values
        self.notion_api_key = os.getenv('NOTION_API_TOKEN') or self.config.get('notion_api_key')
        self.database_id = self.config.get('database_id')
        
        # Load special cases and common words from configuration
        self.special_cases = set(self.config.get('special_cases', []))
        self.common_words = set(self.config.get('common_words', []))
        self.common_words_with_punct = set(self.config.get('common_words_with_punct', []))
        
        # Validate that all required configuration is present
        if not all([self.notion_api_key, self.database_id]):
            raise ValueError("Missing required configuration values: notion_api_key and database_id")
        
        # Set up Notion API request headers
        self.headers = {
            'Authorization': f'Bearer {self.notion_api_key}',
            'Notion-Version': '2022-06-28',
            'Content-Type': 'application/json'
        }
        
        # Initialize optional spaCy NLP support for entity detection
        self.spacy_nlp = None
        if self.config.get('use_spacy', False):
            try:
                import spacy
                self.spacy_nlp = spacy.load("en_core_web_sm")
                logging.info("🧠 spaCy loaded successfully for entity detection")
            except ImportError:
                logging.warning("⚠️  spaCy requested but not available, falling back to heuristics")
        
        # Log successful initialization with key details
        logging.info("✅ Name normalizer initialized")
        logging.info(f"📊 Database ID: {self.database_id}")
        logging.info(f"🔤 Special cases loaded: {len(self.special_cases)}")
        logging.info(f"📝 Common words loaded: {len(self.common_words)}")
        logging.info(f"❓ Common words with punctuation loaded: {len(self.common_words_with_punct)}")
    
    def _load_config(self, config_path: str) -> Dict:
        """
        Load and parse the JSON configuration file.
        
        Args:
            config_path: Path to configuration file
            
        Returns:
            Dictionary containing configuration data
            
        Raises:
            FileNotFoundError: If config file not found
            ValueError: If JSON is invalid
        """
        try:
            with open(config_path, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            # Fallback: try looking in the same directory as this script
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
        """
        Query Notion database for articles created in the last N days.
        
        Args:
            days: Number of days to look back
            
        Returns:
            List of page objects from Notion API
            
        Raises:
            requests.exceptions.RequestException: If API request fails
        """
        pages = []
        
        # Calculate the cutoff date for filtering articles
        filter_date = datetime.utcnow() - timedelta(days=days)
        filter_date_str = filter_date.isoformat() + 'Z'
        
        logging.info(f"🔍 Querying database {self.database_id} for articles created in the last {days} days (since {filter_date_str})")
        
        # Build filter for created_time with descending sort by creation date
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
        
        # Handle pagination through all results
        start_cursor = None
        page_count = 0
        
        while True:
            # Add pagination cursor if we have one from previous request
            if start_cursor:
                filter_body["start_cursor"] = start_cursor
            
            try:
                # Make API request to Notion database
                response = requests.post(
                    f"https://api.notion.com/v1/databases/{self.database_id}/query",
                    headers=self.headers,
                    json=filter_body
                )
                response.raise_for_status()
                
                # Process response data
                data = response.json()
                results = data.get('results', [])
                
                if not results:
                    break  # No more results to process
                
                # Accumulate pages and update counters
                pages.extend(results)
                page_count += len(results)
                logging.debug(f"📥 Retrieved {len(results)} pages (total: {page_count})")
                
                # Check if there are more pages to fetch
                if not data.get('has_more', False):
                    break
                
                # Get cursor for next page
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
        """
        Extract essential information from a Notion page object.
        
        Args:
            page: Notion page object from API
            
        Returns:
            Tuple of (page_id, last_edited_time, article_name) or None if invalid
        """
        page_id = page.get('id', '')
        last_edited_time = page.get('last_edited_time', '')
        
        # Extract article name from the 'article' property (title type)
        properties = page.get('properties', {})
        name_property = properties.get('article', {})
        
        # Validate that article property exists and is of correct type
        if not name_property or name_property.get('type') != 'title':
            logging.warning(f"⚠️  Page {page_id} missing or invalid article property")
            return None
        
        # Extract plain text content from rich text array
        title_content = name_property.get('title', [])
        if not title_content:
            logging.warning(f"⚠️  Page {page_id} has empty article property")
            return None
        
        # Concatenate all text segments and clean up whitespace
        name_text = ''.join([segment.get('plain_text', '') for segment in title_content])
        if not name_text.strip():
            logging.warning(f"⚠️  Page {page_id} has empty article text")
            return None
        
        return page_id, last_edited_time, name_text.strip()
    
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
        
        # Store original tokens for comparison during restoration process
        original_tokens = original_name.split()
        
        # Step 1: Apply basic sentence case - capitalize first character only
        normalized = self._apply_sentence_case(original_name)
        
        # Step 2: Process tokens to restore proper names and handle sentence boundaries
        normalized_tokens = normalized.split()
        result_tokens = []
        i = 0
        
        while i < len(normalized_tokens):
            # Check if we can form a multi-word proper name starting at current position
            multi_word_found = False
            whitelist = self.config.get('proper_name_whitelist', [])
            
            # Look for multi-word proper names in the whitelist
            for proper_name in whitelist:
                if ' ' in proper_name:  # Only process multi-word names
                    proper_words = proper_name.lower().split()
                    proper_length = len(proper_words)
                    
                    # Check if we have enough remaining tokens and they match the proper name
                    if (i + proper_length <= len(normalized_tokens) and 
                        [word.lower() for word in normalized_tokens[i:i + proper_length]] == proper_words):
                        # Additional validation: ensure this is actually a proper name match
                        # and not just common words that happen to appear together
                        if self._is_valid_multi_word_match(normalized_tokens[i:i + proper_length], proper_name):
                            # Found valid multi-word proper name, add it and skip ahead
                            result_tokens.extend(proper_name.split())
                            i += proper_length
                            multi_word_found = True
                            logging.debug(f"🔄 Restored multi-word proper name: {' '.join(normalized_tokens[i:i+proper_length])} -> {proper_name}")
                            break
                        else:
                            logging.debug(f"⚠️  Rejected multi-word match: {' '.join(normalized_tokens[i:i+proper_length])} -> {proper_name} (context doesn't suggest proper name)")
            
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
                
                # Apply token-specific capitalization restoration rules
                restored_token = self._restore_token_capitalization(orig_token, norm_token)
                result_tokens.append(restored_token)
                i += 1
        
        # Step 3: Reconstruct string and normalize whitespace
        result = ' '.join(result_tokens)
        result = re.sub(r'\s+', ' ', result).strip()
        
        return result
    
    def _should_capitalize_token(self, token_index: int, normalized_tokens: List[str], original_tokens: List[str]) -> bool:
        """
        Determine if a token should be capitalized based on sentence boundaries.
        
        Args:
            token_index: Position of token in the sentence
            normalized_tokens: List of normalized tokens
            original_tokens: List of original tokens
            
        Returns:
            True if token should be capitalized, False otherwise
        """
        if token_index == 0:
            return True  # Always capitalize first token of sentence
        
        # Check if previous token ends with sentence-ending punctuation
        prev_token = normalized_tokens[token_index - 1]
        if re.search(r'[.!?]$', prev_token):
            return True
        
        return False
    
    def _apply_sentence_case(self, text: str) -> str:
        """
        Apply basic sentence case: capitalize first letter only.
        
        Args:
            text: Text to apply sentence case to
            
        Returns:
            Text with first letter capitalized
        """
        if not text:
            return text
        
        # Only capitalize the first letter of the entire text
        # Words after proper names should remain lowercase
        if text and text[0].isalpha():
            text = text[0].upper() + text[1:]
        
        return text
    
    def _restore_token_capitalization(self, original_token: str, normalized_token: str) -> str:
        """
        Restore proper capitalization for a specific token based on preservation rules.
        
        Args:
            original_token: Original token with original capitalization
            normalized_token: Token in normalized (sentence case) form
            
        Returns:
            Token with appropriate capitalization applied
        """
        # Rule 1: Preserve ALL CAPS words (2+ chars) - assumed to be acronyms/emphasis
        # This prevents "TEST ALL CAPS" from becoming "Test all caps"
        if len(original_token) >= 2 and original_token.isupper():
            return original_token
        
        # Rule 2: Preserve tokens with special punctuation (e.g., "U.S.A.", "A&B")
        if re.search(r'[.&-]', original_token):
            return original_token
        
        # Rule 3: Preserve the pronoun "I" - always capitalize it
        if original_token.lower() == 'i':
            logging.debug(f"🔤 Preserving pronoun 'I' capitalization")
            return 'I'
        
        # Rule 4: Check against user-defined proper name whitelist from config
        whitelist = self.config.get('proper_name_whitelist', [])
        for proper_name in whitelist:
            if self._is_token_match(original_token, proper_name):
                return proper_name
        
        # Rule 5: Use spaCy NLP to detect person names if available
        if self.spacy_nlp:
            if self._is_person_entity(original_token):
                return original_token
        
        # Rule 6: Handle tokens that contain punctuation (improved logic)
        # This prevents "Karpathy:" from becoming "karpathy:" but still applies sentence case rules
        if re.search(r'[^\w\s]', original_token):
            # Extract the alphabetic part
            alpha_part = re.sub(r'[^\w]', '', original_token)
            if alpha_part:
                # Check if this is a common word that should be lowercased
                alpha_lower = alpha_part.lower()
                if alpha_lower in self.common_words_with_punct:
                    # For common words with punctuation, apply lowercase
                    return self._reconstruct_with_punctuation(alpha_lower, original_token)
                elif len(alpha_part) >= 3 and alpha_part.isupper():
                    # Preserve ALL CAPS words (3+ chars) as acronyms
                    return original_token
                elif len(alpha_part) >= 2 and alpha_part[0].isupper() and alpha_part[1:].islower():
                    # This might be a proper name - check if it's in whitelist or special cases
                    # First check for exact matches
                    exact_match = any(alpha_part.lower() == proper_name.lower() for proper_name in self.config.get('proper_name_whitelist', []))

                    # Then check for partial matches in multi-word names (but not common words)
                    partial_match = False
                    if not exact_match:
                        for proper_name in self.config.get('proper_name_whitelist', []):
                            if ' ' in proper_name:  # Multi-word name
                                proper_words = proper_name.lower().split()
                                if alpha_lower in proper_words and alpha_lower not in self.common_words:
                                    partial_match = True
                                    break

                    if exact_match or partial_match:
                        return original_token
                    else:
                        # Not a known proper name, apply lowercase for sentence case
                        return self._reconstruct_with_punctuation(alpha_lower, original_token)
                else:
                    # For other cases (mixed case, etc.), apply lowercase
                    return self._reconstruct_with_punctuation(alpha_lower, original_token)
        
        # If no preservation rules apply, use the normalized (sentence case) version
        return normalized_token

    def _reconstruct_with_punctuation(self, new_alpha_part: str, original_token: str) -> str:
        """
        Reconstruct a token with punctuation, replacing the alphabetic part.

        Args:
            new_alpha_part: The new alphabetic part to use
            original_token: Original token with punctuation

        Returns:
            Token with new alphabetic part and preserved punctuation
        """
        # Find all non-alphabetic parts (punctuation) in the original token
        parts = re.findall(r'[^\w]+|\w+', original_token)

        # Replace the first alphabetic part with our new version
        result_parts = []
        alpha_replaced = False

        for part in parts:
            if part.isalpha() and not alpha_replaced:
                result_parts.append(new_alpha_part)
                alpha_replaced = True
            else:
                result_parts.append(part)

        return ''.join(result_parts)

    def _is_token_match(self, token: str, proper_name: str) -> bool:
        """
        Check if a token matches a proper name (case-insensitive word boundary match).
        
        Args:
            token: Token to check
            proper_name: Proper name to match against
            
        Returns:
            True if token matches the proper name, False otherwise
        """
        # Convert both to lowercase for case-insensitive comparison
        token_lower = token.lower()
        proper_lower = proper_name.lower()
        
        # Check for exact match first (most common case)
        if token_lower == proper_lower:
            return True
        
        # Check if token is a component of a multi-word proper name
        # BUT only if the token is actually part of the proper name, not just a common word
        # This prevents "law" from matching "Moore's Law" when we're processing "The law of reversed effort"
        # AND prevents "Karpathy" from matching "Andrej Karpathy" when processing "Karpathy: software is changing"
        if ' ' in proper_name:  # Only for multi-word names
            proper_words = proper_lower.split()
            # Only match if the token is a distinctive part of the proper name
            # Avoid matching common words like "law", "future", "work", etc.
            if token_lower in proper_words:
                # Additional check: don't substitute common words that appear in many proper names
                if token_lower not in self.common_words:
                    # CRITICAL FIX: Don't replace individual tokens from multi-word proper names
                    # when they appear in isolation, as they might be in a different context
                    # This prevents "Karpathy" from being replaced when it's just a surname
                    # in a different context like "Karpathy: software is changing"
                    return False
        
        return False
    
    def _is_person_entity(self, token: str) -> bool:
        """
        Check if a token is recognized as a person entity by spaCy NLP.
        
        Args:
            token: Token to check for person entity
            
        Returns:
            True if token is recognized as a person, False otherwise
        """
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
    
    def _is_valid_multi_word_match(self, tokens: List[str], proper_name: str) -> bool:
        """
        Validate if a multi-word match is actually a proper name and not just common words.
        
        Args:
            tokens: List of tokens that potentially match the proper name
            proper_name: Proper name to validate against
            
        Returns:
            True if the match is valid, False otherwise
        """
        # Convert to lowercase for comparison
        token_text = ' '.join([t.lower() for t in tokens])
        proper_lower = proper_name.lower()
        
        # If it's an exact match, it's valid
        if token_text == proper_lower:
            return True
        
        # Check if the tokens contain distinctive words that make it a proper name
        # Common words that appear in many contexts should not trigger substitution
        
        # Count how many distinctive (non-common) words are in the match
        distinctive_words = [word for word in tokens if word.lower() not in self.common_words]
        
        # Require at least 2 distinctive words or the match to be very specific
        if len(distinctive_words) >= 2:
            return True
        
        # Special cases: allow specific proper names even if they contain common words
        # but only if they're distinctive enough
        if proper_name in self.special_cases:
            # These are very specific and should only match when the context is right
            # Check if the surrounding context suggests this is actually the proper name
            return self._context_suggests_proper_name(tokens, proper_name)
        
        return False
    
    def _context_suggests_proper_name(self, tokens: List[str], proper_name: str) -> bool:
        """
        Check if the surrounding context suggests this is actually a proper name reference.
        
        Args:
            tokens: List of tokens to check context for
            proper_name: Proper name to validate context against
            
        Returns:
            True if context suggests this is a proper name, False otherwise
        """
        # For now, be conservative and only allow exact matches for these special cases
        # This prevents "The law of reversed effort" from becoming "The Moore's Law of reversed effort"
        token_text = ' '.join([t.lower() for t in tokens])
        proper_lower = proper_name.lower()
        
        # Only allow substitution if it's an exact match
        return token_text == proper_lower
    
    def _update_page_property(self, page_id: str, property_name: str, new_value: str) -> bool:
        """
        Update a specific property of a Notion page via API.
        
        Args:
            page_id: Notion page ID to update
            property_name: Name of the property to update
            new_value: New value for the property
            
        Returns:
            True if update successful, False otherwise
        """
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
            
            # Make PATCH request to update the page
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
        """
        Main method to process the database and update normalized results.
        
        Args:
            days: Number of days to look back for articles
            
        Returns:
            List of processing results with original and normalized names
        """
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
            
            # Store results for reporting and analysis
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
        
        # Log final processing summary with counts
        logging.info(f"✅ Successfully processed {processed_count} pages")
        logging.info(f"📝 Updated {updated_count} pages with normalized names")
        logging.info(f"✅ Left unchanged {unchanged_count} pages (already normalized)")
        return results


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
    
    args = parser.parse_args()
    
    try:
        # Set up logging first (user must explicitly use --debug if they want detailed output)
        setup_logging(args.debug)
        
        # Test mode: normalize a specific string without database processing
        if args.test:
            # Initialize normalizer for testing
            normalizer = NotionNameNormalizer(args.config)
            
            logging.info(f"🧪 Test mode: normalizing string: '{args.test}'")
            normalized = normalizer._normalize_name(args.test)
            print(f"\nOriginal: {args.test}")
            print(f"Normalized: {normalized}")
            print(f"Changed: {'Yes' if args.test != normalized else 'No'}")
            return
        
        # Initialize normalizer for database processing
        normalizer = NotionNameNormalizer(args.config)
        
        # Process database and update article properties
        results = normalizer.process_database(args.days)
        
        # Log successful completion
        logging.info("✅ Script completed successfully")
        
    except Exception as e:
        logging.error(f"❌ Fatal error: {e}")
        if args.debug:
            logging.exception("Full traceback:")
        sys.exit(1)


if __name__ == '__main__':
    main()