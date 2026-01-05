import json
import os
from pathlib import Path
import pandas as pd
import streamlit as st

# Import history manager functions
from history_manager import get_history_file_path

def load_config():
    """Load configuration from the JSON file in the same directory."""
    config_path = Path(__file__).parent / "linkedin_profiles_data.json"
    if not config_path.exists():
        st.error(f"Config file not found at {config_path}")
        return None

    try:
        with open(config_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        st.error(f"Error loading config: {e}")
        return None

def load_history_data():
    """Load history data from the history Excel file."""
    history_path = get_history_file_path()
    if not history_path or not os.path.exists(history_path):
        return None

    try:
        df = pd.read_excel(history_path)

        # Ensure timestamp is datetime
        if 'timestamp' in df.columns:
            df['timestamp'] = pd.to_datetime(df['timestamp'], errors='coerce')

        # Sort by timestamp descending (most recent first)
        df = df.sort_values('timestamp', ascending=False)

        return df
    except Exception as e:
        st.error(f"Error loading history file: {e}")
        return None

def fuzzy_search_names(df, search_query, max_results=50):
    """
    Perform intelligent fuzzy search on names with multi-word support and relevance ranking.

    Adapted from dataentry.py for history search.
    """
    import re
    import difflib

    if not search_query.strip() or 'name' not in df.columns:
        return df.head(0)

    # Split search query into words and normalize
    search_words = [word.lower().strip() for word in re.split(r'\s+', search_query.strip()) if word.strip()]
    if not search_words:
        return df.head(0)

    results = []

    for idx, row in df.iterrows():
        name = str(row.get('name', '')).lower().strip()
        if not name:
            continue

        # Split name into words for comparison
        name_words = re.split(r'\s+', name)

        # Calculate relevance score using multi-strategy matching
        total_score = 0
        matched_words = 0

        for search_word in search_words:
            best_score = 0

            for name_word in name_words:
                name_word_lower = name_word.lower()

                # Strategy 1: Exact match (highest priority)
                if search_word == name_word_lower:
                    best_score = 100
                    break

                # Strategy 2: Partial substring match
                elif search_word in name_word_lower:
                    score = 80 * (len(search_word) / len(name_word_lower))
                    if score > best_score:
                        best_score = score

                # Strategy 3: Fuzzy matching for typos/similar words
                else:
                    ratio = difflib.SequenceMatcher(None, search_word, name_word_lower).ratio()
                    if ratio > 0.8:  # Only high similarity matches
                        score = 60 * ratio
                        if score > best_score:
                            best_score = score

                    # Strategy 4: Prefix matching for abbreviations
                    if len(search_word) >= 2 and len(name_word_lower) > len(search_word):
                        if name_word_lower.startswith(search_word):
                            score = 70 * (len(search_word) / len(name_word_lower))
                            if score > best_score:
                                best_score = score

            # Accumulate score if we found any match for this search word
            if best_score > 0:
                total_score += best_score
                matched_words += 1

        # Only include results that match at least one search word
        if matched_words > 0:
            # Apply final scoring bonuses
            word_match_bonus = matched_words * 10
            exact_bonus = 20 if any(
                any(search_word == name_word.lower() for name_word in name_words)
                for search_word in search_words
            ) else 0

            final_score = total_score + word_match_bonus + exact_bonus
            results.append((idx, final_score, row))

    # Sort by score (descending) and return top results
    results.sort(key=lambda x: x[1], reverse=True)

    # Extract the rows and return as DataFrame
    if results:
        top_results = results[:max_results]
        return pd.DataFrame([row for _, _, row in top_results])
    else:
        return df.head(0)

def main():
    """Main history viewer function."""

    # Load config
    config = load_config()
    if not config:
        return

    # Load history data
    history_df = load_history_data()
    if history_df is None or history_df.empty:
        st.info("No history data available. Start making changes to contacts to see history here.")
        return

    # Search Section (full width)
    st.subheader("🔍 Search History")

    search_name = st.text_input(
        "Search by name:",
        placeholder="Type a name to search...",
        help="Smart search: supports partial matches, multiple words, and fuzzy matching",
        key="history_search_input"
    )

    show_all = st.checkbox("Show all records", value=False, key="history_show_all")

    # Selection Section (full width)
    st.subheader("📋 Select Contact")

    # Enhanced search with fuzzy matching
    if search_name and not show_all:
        filtered_df = fuzzy_search_names(history_df, search_name)
    elif show_all:
        filtered_df = history_df
    else:
        filtered_df = history_df.head(0)  # Empty dataframe when no search criteria

    if not filtered_df.empty:
        # Get unique names from filtered results
        unique_names = sorted(filtered_df['name'].dropna().unique().tolist())
        contact_options = ["--- Select a contact ---"] + unique_names

        selected_contact = st.selectbox(
            "Choose a contact:",
            options=contact_options,
            help="Select a contact to view their history",
            key="history_contact_selector"
        )

        if selected_contact and selected_contact != "--- Select a contact ---":
            # Filter history for selected contact
            contact_history = history_df[history_df['name'] == selected_contact].copy()

            if not contact_history.empty:
                    st.success(f"📜 Showing history for: {selected_contact}")

                    # Contact info summary
                    latest_record = contact_history.iloc[0]  # Already sorted by timestamp desc

                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Total Changes", len(contact_history))
                    with col2:
                        if pd.notna(latest_record.get('date_contacted')):
                            st.metric("Last Contacted", latest_record['date_contacted'].strftime("%Y-%m-%d"))
                        else:
                            st.metric("Last Contacted", "Never")
                    with col3:
                        if pd.notna(latest_record.get('date_connected')):
                            st.metric("Last Connected", latest_record['date_connected'].strftime("%Y-%m-%d"))
                        else:
                            st.metric("Last Connected", "Never")
                    with col4:
                        if pd.notna(latest_record.get('date_revocation')):
                            st.metric("Last Revoked", latest_record['date_revocation'].strftime("%Y-%m-%d"))
                        else:
                            st.metric("Last Revoked", "Active")

            else:
                st.warning(f"No history found for {selected_contact}")
        else:
            # Show summary when no contact selected
            st.info("Select a contact above to view their detailed history.")

            # Show overall stats
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total History Records", len(history_df))
            with col2:
                unique_contacts = history_df['name'].nunique()
                st.metric("Contacts with History", unique_contacts)
            with col3:
                latest_update = history_df['timestamp'].max()
                if pd.notna(latest_update):
                    st.metric("Latest Update", latest_update.strftime("%Y-%m-%d %H:%M"))

            # Show recent activity
            st.subheader("🕐 Recent Activity")
            recent_df = history_df.head(10)[['timestamp', 'action', 'name', 'company']].copy()
            # Convert to datetime and format valid timestamps only
            valid_timestamps = pd.to_datetime(recent_df['timestamp'], errors='coerce')
            mask = valid_timestamps.notna()
            recent_df.loc[mask, 'timestamp'] = valid_timestamps.loc[mask].dt.strftime("%Y-%m-%d %H:%M")
            st.dataframe(recent_df, width='stretch')

    else:
        if search_name:
            st.warning("No contacts found matching your search.")
        st.info("Use the search above or check 'Show all records' to view contact history.")

    # Change History Section (full width)
    if 'selected_contact' in locals() and selected_contact and selected_contact != "--- Select a contact ---":
        contact_history = history_df[history_df['name'] == selected_contact].copy()

        if not contact_history.empty:
            st.subheader("📅 Change History")

            # Prepare display columns
            display_cols = ['timestamp', 'action', 'date_contacted', 'date_connected', 'date_revocation', 'company', 'job_title']
            available_cols = [col for col in display_cols if col in contact_history.columns]

            display_df = contact_history[available_cols].copy()

            # Format timestamps and dates
            if 'timestamp' in display_df.columns:
                # Convert to datetime and format valid timestamps only
                valid_timestamps = pd.to_datetime(display_df['timestamp'], errors='coerce')
                mask = valid_timestamps.notna()
                display_df.loc[mask, 'timestamp'] = valid_timestamps.loc[mask].dt.strftime("%Y-%m-%d %H:%M:%S")

            date_cols = ['date_contacted', 'date_connected', 'date_revocation']
            for col in date_cols:
                if col in display_df.columns:
                    # Convert to datetime and format valid dates only
                    valid_dates = pd.to_datetime(display_df[col], errors='coerce')
                    mask = valid_dates.notna()
                    # Ensure column is object dtype before assignment to avoid dtype warnings
                    display_df[col] = display_df[col].astype(object)
                    display_df.loc[mask, col] = valid_dates.loc[mask].dt.strftime("%Y-%m-%d")

            # Rename columns for better readability
            column_names = {
                'timestamp': 'When Changed',
                'action': 'Action Type',
                'date_contacted': 'Contact Date',
                'date_connected': 'Connected Date',
                'date_revocation': 'Revoked Date',
                'company': 'Company',
                'job_title': 'Job Title'
            }

            display_df = display_df.rename(columns=column_names)
            # Reset index so styling compares rows in displayed order
            display_df = display_df.reset_index(drop=True)

            # Function to highlight cells that changed from previous row
            def highlight_changes(data):
                """
                Highlight cells that have changed compared to the previous row.
                Excludes the 'When Changed' column since timestamps are always different.
                Returns a DataFrame of styles with blue background for changed cells.
                """
                # Create a DataFrame to store styles with aligned positional index
                styles = pd.DataFrame('', index=data.index, columns=data.columns)

                # Compare each row to the next (older) row since data is sorted desc
                for idx in range(0, len(data) - 1):
                    curr_row = data.iloc[idx]
                    older_row = data.iloc[idx + 1]

                    # Compare each column except 'When Changed' (timestamp)
                    for col in data.columns:
                        if col == 'When Changed':
                            continue  # Skip timestamp column

                        curr_val = curr_row[col]
                        older_val = older_row[col]

                        # Check if values are different (handle NaN values)
                        if pd.isna(curr_val) and pd.isna(older_val):
                            continue  # Both NaN, no change
                        elif pd.isna(curr_val) != pd.isna(older_val):
                            # One is NaN and other isn't - this is a change
                            styles.iloc[idx, styles.columns.get_loc(col)] = 'background-color: rgba(30, 136, 229, 0.3);'
                        elif str(curr_val).strip() != str(older_val).strip():
                            # Values are different (as strings for comparison)
                            styles.iloc[idx, styles.columns.get_loc(col)] = 'background-color: rgba(30, 136, 229, 0.3);'

                return styles

            # Apply highlighting and display the styled table
            styled_df = display_df.style.apply(highlight_changes, axis=None)
            st.dataframe(styled_df, width='stretch')