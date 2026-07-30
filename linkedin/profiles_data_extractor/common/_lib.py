"""
Shared search/formatting/color helpers for the LinkedIn profiles data extractor.

All Streamlit pages inside this package (``dataentry``, ``reachout``,
``history``, ``dashboard``) import from here so the fuzzy-search algorithm,
Excel formatting glue, and color-gradient generator each live in exactly one
place instead of being copy-pasted per page.
"""

import difflib
import logging
import re

import pandas as pd
import streamlit as st

from excel_format_manager import convert_url_columns_to_hyperlinks, apply_format_from_json

log = logging.getLogger(__name__)

DEFAULT_SEARCH_COLUMNS = ('name', 'job_title', 'location')


def fuzzy_search_profiles(df, search_query, columns=DEFAULT_SEARCH_COLUMNS, max_results=100):
    """
    Perform intelligent fuzzy search across one or more profile fields with
    multi-word support and relevance ranking.

    This function implements a sophisticated search algorithm that:
    - Splits search queries into multiple words
    - Matches each word against the combined text of ``columns`` using
      multiple strategies
    - Ranks results by relevance score based on match quality and coverage
    - Supports partial matches, fuzzy matching, and prefix matching
    - Combines scores from all matching fields for comprehensive results

    Args:
        df (pandas.DataFrame): DataFrame containing profile data to search
        search_query (str): Search string that can contain multiple words separated by spaces
        columns (tuple[str, ...]): Columns to search across, combined into one text blob
        max_results (int): Maximum number of results to return (default: 100)

    Returns:
        pandas.DataFrame: Filtered DataFrame sorted by relevance score (highest first)
                         Empty DataFrame if no matches found

    Examples:
        >>> df = pd.DataFrame({'name': ['Ana Izquierdo'], 'job_title': ['Software Engineer'], 'location': ['Madrid']})
        >>> fuzzy_search_profiles(df, 'ana engineer')  # Returns profile with matches in name and job_title
        >>> fuzzy_search_profiles(df, 'ana izq', columns=('name',))  # Returns Ana Izquierdo first
    """
    if not search_query.strip():
        return df.head(0)

    # Check if we have at least one searchable column
    available_columns = [col for col in columns if col in df.columns]
    if not available_columns:
        return df.head(0)

    # Split search query into words and normalize
    search_words = [word.lower().strip() for word in re.split(r'\s+', search_query.strip()) if word.strip()]
    if not search_words:
        return df.head(0)

    results = []

    for idx, row in df.iterrows():
        # Combine all searchable field content for this profile
        profile_text = ' '.join([
            str(row.get(col, '')).lower().strip()
            for col in available_columns
        ]).strip()

        if not profile_text:
            continue

        # Split combined text into words for comparison
        profile_words = re.split(r'\s+', profile_text)

        # Calculate relevance score using multi-strategy matching
        # Each search word is matched against the combined profile text
        total_score = 0
        matched_words = 0

        for search_word in search_words:
            best_score = 0

            for profile_word in profile_words:
                profile_word_lower = profile_word.lower()

                # Strategy 1: Exact match (highest priority)
                # "ana" exactly matches "ana" → 100 points
                if search_word == profile_word_lower:
                    best_score = 100
                    break

                # Strategy 2: Partial substring match
                # "izq" is substring of "izquierdo" → 80 points scaled by length ratio
                elif search_word in profile_word_lower:
                    score = 80 * (len(search_word) / len(profile_word_lower))
                    if score > best_score:
                        best_score = score

                # Strategy 3: Fuzzy matching for typos/similar words
                else:
                    # Use difflib for sequence similarity (handles typos like "izq" ≈ "izqu")
                    ratio = difflib.SequenceMatcher(None, search_word, profile_word_lower).ratio()
                    if ratio > 0.8:  # Only high similarity matches
                        score = 60 * ratio
                        if score > best_score:
                            best_score = score

                    # Strategy 4: Prefix matching for abbreviations
                    # "ana" matches start of "ana maría" → 70 points scaled by coverage
                    if len(search_word) >= 2 and len(profile_word_lower) > len(search_word):
                        if profile_word_lower.startswith(search_word):
                            score = 70 * (len(search_word) / len(profile_word_lower))
                            if score > best_score:
                                best_score = score

            # Accumulate score if we found any match for this search word
            if best_score > 0:
                total_score += best_score
                matched_words += 1

        # Only include results that match at least one search word
        if matched_words > 0:
            # Apply final scoring bonuses for better ranking:

            # Bonus for matching multiple search words (encourages comprehensive matches)
            word_match_bonus = matched_words * 10

            # Bonus for having at least one exact word match (prioritizes precision)
            exact_bonus = 20 if any(
                any(search_word == profile_word.lower() for profile_word in profile_words)
                for search_word in search_words
            ) else 0

            final_score = total_score + word_match_bonus + exact_bonus
            results.append((idx, final_score, row))

    # Sort by score (descending) and return top results
    results.sort(key=lambda x: x[1], reverse=True)

    # Extract the rows and return as DataFrame
    if results:
        indices = [idx for idx, score, row in results[:max_results]]
        return df.loc[indices].copy()
    else:
        return df.head(0)


def fuzzy_search_names(df, search_query, max_results=50):
    """
    Perform intelligent fuzzy search on the ``name`` column only.

    Thin wrapper around :func:`fuzzy_search_profiles` restricted to the
    ``name`` column — see that function for the matching strategy details.

    Args:
        df (pandas.DataFrame): DataFrame containing a 'name' column to search
        search_query (str): Search string that can contain multiple words separated by spaces
        max_results (int): Maximum number of results to return (default: 50)

    Returns:
        pandas.DataFrame: Filtered DataFrame sorted by relevance score (highest first)
                         Empty DataFrame if no matches found

    Examples:
        >>> df = pd.DataFrame({'name': ['Ana Izquierdo', 'Juan Pérez', 'Ana García']})
        >>> fuzzy_search_names(df, 'ana izq')  # Returns Ana Izquierdo first
        >>> fuzzy_search_names(df, 'ana')      # Returns all Ana* names
    """
    return fuzzy_search_profiles(df, search_query, columns=('name',), max_results=max_results)


def save_to_excel(df, file_path):
    """Save DataFrame to Excel file."""
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)
        return True
    except Exception as e:
        st.error(f"Error saving to Excel: {e}")
        return False


def apply_excel_formatting(excel_path, json_path):
    """Apply Excel formatting: convert URLs to hyperlinks and apply format from JSON."""
    try:
        log.info("🔗 Converting URL columns to hyperlinks...")
        # Convert URL columns to hyperlinks
        url_success = convert_url_columns_to_hyperlinks(excel_path)
        if url_success:
            log.info("✅ URL hyperlinks conversion completed")
        else:
            log.error("❌ URL hyperlinks conversion failed")

        log.info("🎨 Applying formatting from JSON specification...")
        # Apply formatting from JSON
        format_success = apply_format_from_json(excel_path, json_path)
        if format_success:
            log.info("✅ JSON formatting applied successfully")
        else:
            log.error("❌ JSON formatting application failed")

        overall_success = url_success and format_success
        log.info("🎯 Excel formatting process %s", 'completed successfully' if overall_success else 'failed')
        return overall_success
    except Exception as e:
        log.error("❌ Error applying Excel formatting: %s", e)
        return False


def generate_color_gradient(start_hex, end_hex, n):
    """Generate a gradient of n colors between start_hex and end_hex."""
    if n < 1:
        return []
    if n == 1:
        return [start_hex]

    def hex_to_rgb(h):
        return tuple(int(h.lstrip('#')[i:i + 2], 16) for i in (0, 2, 4))

    start_rgb = hex_to_rgb(start_hex)
    end_rgb = hex_to_rgb(end_hex)

    colors = []
    for i in range(n):
        ratio = i / (n - 1)
        rgb = tuple(int(start_rgb[j] + (end_rgb[j] - start_rgb[j]) * ratio) for j in range(3))
        colors.append('#{:02x}{:02x}{:02x}'.format(*rgb))

    return colors
