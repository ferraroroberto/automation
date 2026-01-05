import json
import os
from pathlib import Path
from datetime import datetime
import re
import difflib

import pandas as pd
import streamlit as st
import plotly.express as px

# Import Excel formatting functions
from excel_format_manager import convert_url_columns_to_hyperlinks, apply_format_from_json
from history_manager import log_history

def generate_color_gradient(start_hex, end_hex, n):
    """Generate a gradient of n colors between start_hex and end_hex."""
    if n < 1: return []
    if n == 1: return [start_hex]

    def hex_to_rgb(h):
        return tuple(int(h.lstrip('#')[i:i+2], 16) for i in (0, 2, 4))

    start_rgb = hex_to_rgb(start_hex)
    end_rgb = hex_to_rgb(end_hex)

    colors = []
    for i in range(n):
        ratio = i / (n - 1)
        rgb = tuple(int(start_rgb[j] + (end_rgb[j] - start_rgb[j]) * ratio) for j in range(3))
        colors.append('#{:02x}{:02x}{:02x}'.format(*rgb))

    return colors

def fuzzy_search_profiles(df, search_query, max_results=100):
    """
    Perform intelligent fuzzy search across multiple profile fields with multi-word support and relevance ranking.

    This function implements a sophisticated search algorithm that:
    - Splits search queries into multiple words
    - Matches each word against profile fields (name, job_title, location) using multiple strategies
    - Ranks results by relevance score based on match quality and coverage
    - Supports partial matches, fuzzy matching, and prefix matching
    - Combines scores from all matching fields for comprehensive results

    Args:
        df (pandas.DataFrame): DataFrame containing profile data to search
        search_query (str): Search string that can contain multiple words separated by spaces
        max_results (int): Maximum number of results to return (default: 100)

    Returns:
        pandas.DataFrame: Filtered DataFrame sorted by relevance score (highest first)
                         Empty DataFrame if no matches found

    Examples:
        >>> df = pd.DataFrame({'name': ['Ana Izquierdo'], 'job_title': ['Software Engineer'], 'location': ['Madrid']})
        >>> fuzzy_search_profiles(df, 'ana engineer')  # Returns profile with matches in name and job_title
        >>> fuzzy_search_profiles(df, 'madrid')        # Returns profile with location match
    """
    if not search_query.strip():
        return df.head(0)

    # Check if we have at least one searchable column
    searchable_columns = ['name', 'job_title', 'location']
    available_columns = [col for col in searchable_columns if col in df.columns]
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
            best_match_type = 'none'

            for profile_word in profile_words:
                profile_word_lower = profile_word.lower()

                # Strategy 1: Exact match (highest priority)
                # "ana" exactly matches "ana" → 100 points
                if search_word == profile_word_lower:
                    best_score = 100
                    best_match_type = 'exact'
                    break

                # Strategy 2: Partial substring match
                # "izq" is substring of "izquierdo" → 80 points scaled by length ratio
                elif search_word in profile_word_lower:
                    score = 80 * (len(search_word) / len(profile_word_lower))
                    if score > best_score:
                        best_score = score
                        best_match_type = 'partial'

                # Strategy 3: Fuzzy matching for typos/similar words
                else:
                    # Use difflib for sequence similarity (handles typos like "izq" ≈ "izqu")
                    ratio = difflib.SequenceMatcher(None, search_word, profile_word_lower).ratio()
                    if ratio > 0.8:  # Only high similarity matches
                        score = 60 * ratio
                        if score > best_score:
                            best_score = score
                            best_match_type = 'fuzzy'

                    # Strategy 4: Prefix matching for abbreviations
                    # "ana" matches start of "ana maría" → 70 points scaled by coverage
                    if len(search_word) >= 2 and len(profile_word_lower) > len(search_word):
                        if profile_word_lower.startswith(search_word):
                            score = 70 * (len(search_word) / len(profile_word_lower))
                            if score > best_score:
                                best_score = score
                                best_match_type = 'prefix'

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

def load_excel_data(file_path):
    """Load data from the Excel file."""
    if not os.path.exists(file_path):
        st.error(f"Data file not found at: {file_path}")
        return None

    try:
        df = pd.read_excel(file_path)

        # Ensure date columns are datetime
        date_columns = ['day', 'date connected', 'revocation_date']
        for col in date_columns:
            if col not in df.columns:
                df[col] = pd.NaT  # Create missing column
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')

        return df
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return None

def save_to_excel(df, file_path):
    """Save DataFrame to Excel file."""
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)
        return True
    except Exception as e:
        st.error(f"Error saving to Excel: {e}")
        return False

def update_existing_record(df, original_name, record_data):
    """Update an existing record by original name."""
    # Find the record by original name (case-insensitive)
    existing_mask = df['name'].str.lower() == original_name.lower() if 'name' in df.columns else pd.Series([False] * len(df))

    if existing_mask.any():
        # Update existing record
        idx = df[existing_mask].index[0]
        for key, value in record_data.items():
            if key in df.columns:
                df.at[idx, key] = value
        return df, "updated"
    else:
        return df, "not_found"

def apply_excel_formatting(excel_path, json_path):
    """Apply Excel formatting: convert URLs to hyperlinks and apply format from JSON."""
    try:
        print("🔗 Converting URL columns to hyperlinks...")
        # Convert URL columns to hyperlinks
        url_success = convert_url_columns_to_hyperlinks(excel_path)
        if url_success:
            print("✅ URL hyperlinks conversion completed")
        else:
            print("❌ URL hyperlinks conversion failed")

        print("🎨 Applying formatting from JSON specification...")
        # Apply formatting from JSON
        format_success = apply_format_from_json(excel_path, json_path)
        if format_success:
            print("✅ JSON formatting applied successfully")
        else:
            print("❌ JSON formatting application failed")

        overall_success = url_success and format_success
        print(f"🎯 Excel formatting process {'completed successfully' if overall_success else 'failed'}")
        return overall_success
    except Exception as e:
        print(f"❌ Error applying Excel formatting: {e}")
        return False

def main(df_filtered, df_all):
    """Main reachout function for managing uncontacted profiles."""

    # Load config to get data path
    config = load_config()
    if not config:
        return
    data_path = config.get("destination_file")
    if not data_path:
        st.error("No 'destination_file' specified in config.")
        return

    # Filter to only show uncontacted profiles (day is null) from the already filtered data
    # Create view selector
    view_mode = st.radio("View Mode", ["Uncontacted Profiles", "Revoked Contacts"], horizontal=True)

    if view_mode == "Uncontacted Profiles":
        uncontacted_df = df_filtered[df_filtered['day'].isna()].copy() if 'day' in df_filtered.columns else df_filtered.copy()
        
        # Filter out records that are actually revoked (have a revocation date)
        if 'revocation_date' in uncontacted_df.columns:
             # Just to be safe, though they shouldn't have a 'day' anyway
             pass
             
    else: # Revoked Contacts
        if 'revocation_date' in df_filtered.columns and 'day' in df_filtered.columns:
            # Filter where revocation_date is present AND revocation_date > day
            # Note: We must handle cases where 'day' might be NaT or populated.
            # Logic: A contact is "revoked" if it HAS been contacted (day is not null) 
            # AND it has a revocation date that is AFTER the contact date.
            
            mask = (df_filtered['day'].notna()) & (df_filtered['revocation_date'].notna()) & (df_filtered['revocation_date'] > df_filtered['day'])
            uncontacted_df = df_filtered[mask].copy()
            
            # Aging Filters for Revoked Contacts
            st.write("🕒 **Aging Filter** (Time since revocation)")
            col_age1, col_age2, col_age3, col_age4 = st.columns(4)
            
            age_filter = st.radio(
                "Show profiles revoked at least:",
                ["All Revoked", "4 Weeks Ago", "8 Weeks Ago", "12 Weeks Ago"],
                horizontal=True,
                label_visibility="collapsed"
            )
            
            if age_filter != "All Revoked":
                now = pd.Timestamp.now()
                weeks = 0
                if age_filter == "4 Weeks Ago": weeks = 4
                elif age_filter == "8 Weeks Ago": weeks = 8
                elif age_filter == "12 Weeks Ago": weeks = 12
                
                cutoff_date = now - pd.Timedelta(weeks=weeks)
                uncontacted_df = uncontacted_df[uncontacted_df['revocation_date'] <= cutoff_date]
                
        else:
            uncontacted_df = pd.DataFrame()
            st.warning("Revocation date column missing from data.")

    if uncontacted_df.empty:
        if view_mode == "Uncontacted Profiles":
            st.info("🎉 All profiles have been contacted! No reachout targets remaining.")
        else:
            st.info("No revoked contacts found matching criteria.")
        return

    # Add horizontal bar chart showing uncontacted people by company at the top
    chart_title = "Uncontacted Profiles" if view_mode == "Uncontacted Profiles" else "Revoked Contacts"
    st.subheader(f"📈 {chart_title} by Company")

    # Create stacked bar chart by company and search_type
    if 'company' in uncontacted_df.columns and 'search_type' in uncontacted_df.columns:
        # Filter out empty/null company names
        filtered_df = uncontacted_df[uncontacted_df['company'].notna() & (uncontacted_df['company'] != '')].copy()

        if not filtered_df.empty:
            # Create cross-tabulation of company vs search_type
            company_search_counts = pd.crosstab(filtered_df['company'], filtered_df['search_type'])

            # Sort companies by total count (ascending)
            company_totals = company_search_counts.sum(axis=1).sort_values(ascending=True)
            company_search_counts = company_search_counts.loc[company_totals.index]

            # Get unique search types and create color gradient
            search_types = company_search_counts.columns.tolist()
            colors = generate_color_gradient('#0B65C3', '#808080', len(search_types))

            # Create stacked horizontal bar chart
            fig = px.bar(
                company_search_counts,
                orientation='h',
                title=f'Number of {chart_title} by Company and Search Type',
                labels={'value': f'Number of {chart_title}', 'company': 'Company'},
                color_discrete_map={search_type: colors[i] for i, search_type in enumerate(search_types)}
            )

            # Customize layout
            fig.update_layout(
                xaxis_title=f"Number of {chart_title}",
                yaxis_title="Company",
                showlegend=True,
                legend_title="Search Type",
                height=max(400, len(company_search_counts) * 35),  # Dynamic height based on number of companies
                barmode='stack'  # Stack the bars
            )

            st.plotly_chart(fig, width='stretch')
        else:
            st.info("No company data available for chart.")
    else:
        st.warning("Required columns (company and search_type) not found in data.")

    st.markdown("---")

    # Search section for uncontacted profiles
    col_search, col_filter = st.columns([2, 1])

    with col_search:
        search_uncontacted = st.text_input(
            f"🔍 Search {view_mode.lower()} (name, job title, or location):",
            placeholder="Type a name, job title, or location to search...",
            help="Smart search: supports partial matches, multiple words, and fuzzy matching (e.g., 'ana izq' finds 'ana izquierdo')",
            key="search_uncontacted"
        )

    with col_filter:
        # Auto-uncheck "show all" when there's search input
        if 'show_all_uncontacted' not in st.session_state:
            st.session_state.show_all_uncontacted = True

        # If search input has content and "show all" is checked, uncheck it
        if search_uncontacted.strip() and st.session_state.show_all_uncontacted:
            st.session_state.show_all_uncontacted = False

        show_all_uncontacted = st.checkbox(f"Show all {view_mode.lower()}", value=st.session_state.show_all_uncontacted, key="show_all_uncontacted")

    # Apply search filtering to uncontacted profiles
    if search_uncontacted and not show_all_uncontacted:
        # Use intelligent fuzzy search that handles partial matches, multiple words,
        # and similarity matching across name, job_title, and location
        filtered_uncontacted_df = fuzzy_search_profiles(uncontacted_df, search_uncontacted)
    elif show_all_uncontacted:
        filtered_uncontacted_df = uncontacted_df
    else:
        filtered_uncontacted_df = uncontacted_df.head(0)  # Empty dataframe when no search criteria

    # Display filtered count
    if search_uncontacted or show_all_uncontacted:
        st.subheader(f"📋 {chart_title} ({len(filtered_uncontacted_df)} filtered from {len(uncontacted_df)} total)")
    else:
        st.subheader(f"📋 {chart_title} ({len(uncontacted_df)} total)")

    # Initialize session state for selected record
    if 'reachout_selected_record' not in st.session_state:
        st.session_state.reachout_selected_record = None
    if 'reachout_original_name' not in st.session_state:
        st.session_state.reachout_original_name = None
    if 'reachout_editing' not in st.session_state:
        st.session_state.reachout_editing = False

    if filtered_uncontacted_df.empty:
        if search_uncontacted:
            st.info(f"No {view_mode.lower()} found matching your search.")
        else:
            st.info("No profiles match the selected filters.")
        return

    # Create two columns layout
    col_list, col_edit = st.columns([1, 1])

    with col_list:
        st.subheader("👥 Select Profile")

        # Create selectable list
        for idx, row in filtered_uncontacted_df.iterrows():
            name = row.get('name', 'Unknown')
            job_title = row.get('job_title', '')
            company = row.get('company', '')
            location = row.get('location', '')

            # Create display text
            display_text = f"**{name}**"
            if job_title:
                display_text += f" - {job_title}"
            if company:
                display_text += f" at {company}"
            if location:
                display_text += f" ({location})"

            # Check if this record is currently selected
            is_selected = (st.session_state.reachout_selected_record and
                          st.session_state.reachout_selected_record.get('name') == name)

            # Create clickable button for each record
            if st.button(display_text, key=f"select_{idx}", width='stretch',
                        help="Click to edit this profile"):
                # Store selected record data
                st.session_state.reachout_selected_record = {
                    'name': name,
                    'day': row.get('day', ''),
                    'job_title': job_title,
                    'follows_from': row.get('follows_from', ''),
                    'company': company,
                    'location': location,
                    'search_type': row.get('search_type', ''),
                    'url': row.get('url', ''),
                    'reach out type': row.get('reach out type', ''),
                    'revocation_date': row.get('revocation_date', None)
                }
                st.session_state.reachout_original_name = name
                st.session_state.reachout_editing = True
                st.rerun()

    with col_edit:
        # Header with Edit Profile and Open Profile link
        col_header, col_link = st.columns([3, 1])
        with col_header:
            st.subheader("✏️ Edit Profile")
        with col_link:
            if st.session_state.reachout_editing and st.session_state.reachout_selected_record:
                profile_url = st.session_state.reachout_selected_record.get('url', '')
                if profile_url:
                    st.markdown(f'<a href="{profile_url}" target="_blank" style="text-decoration: none;"><button style="background-color: #0077b5; color: white; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer;">🔗 Open Profile</button></a>', unsafe_allow_html=True)

        if st.session_state.reachout_editing and st.session_state.reachout_selected_record:
            record = st.session_state.reachout_selected_record

            with st.form("reachout_edit_form"):


                # Editable fields
                name = st.text_input(
                    "Name *",
                    value=record.get('name', ''),
                    help="Full name of the LinkedIn profile"
                )

                # Handle day field - convert to date if it's not None/NaT
                day_value = record.get('day', '')
                if pd.notna(day_value) and day_value != '':
                    try:
                        # Convert to datetime.date if it's a pandas Timestamp
                        if hasattr(day_value, 'date'):
                            day_value = day_value.date()
                        elif isinstance(day_value, str):
                            day_value = pd.to_datetime(day_value).date()
                    except:
                        day_value = None
                else:
                    day_value = None

                day = st.date_input(
                    "Day Contacted",
                    value=day_value,
                    help="Date when this profile was contacted (leave empty if not contacted yet)"
                )
                
                # Show revocation info if present
                revocation_val = record.get('revocation_date')
                if pd.notna(revocation_val):
                    rev_str = revocation_val.strftime('%Y-%m-%d') if hasattr(revocation_val, 'strftime') else str(revocation_val)
                    st.warning(f"⚠️ This contact was revoked on {rev_str}. Updating the 'Day Contacted' will clear the revocation date.")

                # Get distinct reach out types from the full dataframe
                reachout_types = ['']  # Start with empty option
                if 'reach out type' in df_all.columns:
                    distinct_types = df_all['reach out type'].dropna().unique().tolist()
                    reachout_types.extend(sorted(distinct_types))

                reach_out_type = st.selectbox(
                    "Reachout Type",
                    options=reachout_types,
                    index=reachout_types.index(record.get('reach out type', '')) if record.get('reach out type', '') in reachout_types else 0,
                    help="Type of reachout made to this profile"
                )

                search_type = st.text_input(
                    "Search Type",
                    value=record.get('search_type', ''),
                    help="How this profile was found (e.g., keyword search, mutual connections, etc.)"
                )

                job_title = st.text_input(
                    "Job Title",
                    value=record.get('job_title', ''),
                    help="Current job position"
                )

                follows_from = st.text_input(
                    "Follows From",
                    value=record.get('follows_from', ''),
                    help="How you connected with this person"
                )

                company = st.text_input(
                    "Company",
                    value=record.get('company', ''),
                    help="Company name"
                )

                location = st.text_input(
                    "Location",
                    value=record.get('location', ''),
                    help="Location/city"
                )

                url = st.text_input(
                    "Profile URL",
                    value=record.get('url', ''),
                    help="LinkedIn profile URL"
                )

                submitted = st.form_submit_button("💾 Save Changes")

                if submitted:
                    if not name.strip():
                        st.error("❌ Name is required!")
                        return

                    # Prepare record data
                    final_day = pd.Timestamp(day) if day else pd.NaT
                    
                    # Reset Logic: If contact date is updated (and not just cleared), reset revocation date
                    original_day = st.session_state.reachout_selected_record.get('day')
                    # Handle NaT comparison
                    original_day_ts = pd.Timestamp(original_day) if pd.notna(original_day) else pd.NaT
                    
                    # Check if day has changed to a new valid date (re-contact logic)
                    revocation_update = st.session_state.reachout_selected_record.get('revocation_date')
                    if day and (pd.isna(original_day_ts) or final_day != original_day_ts):
                         revocation_update = None
                         # We don't need to show a message here as the save success message is enough, 
                         # but we ensure the dict sends None for revocation_date

                    record_data = {
                        'name': name.strip(),
                        'day': final_day,
                        'job_title': job_title.strip(),
                        'follows_from': follows_from.strip(),
                        'company': company.strip(),
                        'location': location.strip(),
                        'search_type': search_type.strip(),
                        'url': url.strip(),
                        'reach out type': reach_out_type.strip() if reach_out_type else '',
                        'revocation_date': revocation_update
                    }

                    # Update existing record
                    updated_df, action = update_existing_record(df_all, st.session_state.reachout_original_name, record_data)

                    if action == "not_found":
                        st.error("❌ Original record not found. It may have been deleted.")
                        return

                    # Save to Excel
                    if save_to_excel(updated_df, data_path):
                        st.success("✅ Profile updated successfully!")
                        
                        # Log history
                        log_history(record_data, action="update")

                        # Automatically apply Excel formatting after saving
                        json_path = Path(__file__).parent / "excel_format_spec.json"
                        if json_path.exists():
                            print("🎨 Applying Excel formatting...")
                            try:
                                if apply_excel_formatting(data_path, str(json_path)):
                                    print("✅ Excel formatting applied successfully!")
                                else:
                                    print("❌ Failed to apply Excel formatting")
                            except Exception as e:
                                print(f"❌ Error applying Excel formatting: {e}")
                        else:
                            print("⚠️ Format specification file not found. Formatting not applied.")

                        # Clear selection and rerun
                        st.session_state.reachout_selected_record = None
                        st.session_state.reachout_original_name = None
                        st.session_state.reachout_editing = False
                        st.rerun()
                    else:
                        st.error("❌ Failed to save changes!")

            # Delete Profile button (red, similar to Open Profile button)
            st.markdown("---")
            col_delete, col_spacer = st.columns([1, 3])
            with col_delete:
                if st.button("🗑️ Delete Profile", key="delete_profile", type="secondary",
                           help="Permanently delete this profile from the database"):
                    # Confirm deletion with a dialog-like approach
                    confirm_key = f"confirm_delete_{st.session_state.reachout_original_name}"
                    if confirm_key not in st.session_state:
                        st.session_state[confirm_key] = False

                    if not st.session_state[confirm_key]:
                        st.warning(f"⚠️ Are you sure you want to delete '{st.session_state.reachout_original_name}'? This action cannot be undone!")
                        if st.button("✅ Yes, Delete", key=f"confirm_yes_{st.session_state.reachout_original_name}"):
                            st.session_state[confirm_key] = True
                            st.rerun()
                        if st.button("❌ Cancel", key=f"confirm_no_{st.session_state.reachout_original_name}"):
                            del st.session_state[confirm_key]
                            st.rerun()
                    else:
                        # Perform deletion
                        try:
                            # Find and remove the record
                            name_to_delete = st.session_state.reachout_original_name
                            mask = df_all['name'].str.lower() == name_to_delete.lower() if 'name' in df_all.columns else pd.Series([False] * len(df_all))

                            if mask.any():
                                # Get record data for history before deleting
                                record_to_delete = df_all[mask].iloc[0].to_dict()

                                # Remove the record
                                df_all = df_all[~mask].copy()

                                # Save to Excel
                                if save_to_excel(df_all, data_path):
                                    st.success(f"✅ Profile '{name_to_delete}' deleted successfully!")
                                    
                                    # Log history
                                    log_history(record_to_delete, action="delete")

                                    # Automatically apply Excel formatting after saving
                                    json_path = Path(__file__).parent / "excel_format_spec.json"
                                    if json_path.exists():
                                        print("🎨 Applying Excel formatting...")
                                        try:
                                            if apply_excel_formatting(data_path, str(json_path)):
                                                print("✅ Excel formatting applied successfully!")
                                            else:
                                                print("❌ Failed to apply Excel formatting")
                                        except Exception as e:
                                            print(f"❌ Error applying Excel formatting: {e}")
                                    else:
                                        print("⚠️ Format specification file not found. Formatting not applied.")

                                    # Clear session state
                                    st.session_state.reachout_selected_record = None
                                    st.session_state.reachout_original_name = None
                                    st.session_state.reachout_editing = False
                                    if confirm_key in st.session_state:
                                        del st.session_state[confirm_key]
                                    st.rerun()
                                else:
                                    st.error("❌ Failed to save changes after deletion!")
                            else:
                                st.error("❌ Profile not found for deletion.")
                        except Exception as e:
                            st.error(f"❌ Error deleting profile: {str(e)}")
        else:
            st.info("👆 Click on a profile from the list to edit its information.")

    # Show summary table at the bottom
    st.markdown("---")
    st.subheader("📊 Reachout Summary")

    if not filtered_uncontacted_df.empty:
        # Display summary table with the requested columns
        display_cols = ['name', 'job_title', 'follows_from', 'company', 'location']
        available_cols = [col for col in filtered_uncontacted_df.columns]

        if available_cols:
            summary_df = filtered_uncontacted_df[available_cols].copy()
            st.dataframe(summary_df, width='stretch')
        else:
            st.warning("Required columns not found in data.")

if __name__ == "__main__":
    main()