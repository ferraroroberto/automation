from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st

from history_manager import log_history
from loaders import load_config, load_excel_data
from _lib import fuzzy_search_names, save_to_excel, apply_excel_formatting

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

def main():
    """
    Main Streamlit application for LinkedIn profile data entry and editing.

    This application provides a user-friendly interface for:
    - Searching LinkedIn profiles using intelligent fuzzy matching
    - Automatically selecting the first relevant result
    - Editing profile data (name, connection date, response status, chat URLs)
    - Applying Excel formatting automatically after saves

    Features:
    - Multi-word fuzzy search with relevance ranking
    - Auto-selection of first search result for faster workflow
    - Real-time data validation and error handling
    - Excel formatting with hyperlinks and custom styling
    - Session state management for seamless user experience
    """

    # Load config
    config = load_config()
    if not config:
        return

    data_path = config.get("destination_file")
    if not data_path:
        st.error("No 'destination_file' specified in config.")
        return

    # Load current data
    df = load_excel_data(data_path)
    if df is None:
        return

    # Initialize session state for selected record
    if 'selected_record' not in st.session_state:
        st.session_state.selected_record = None
    if 'original_name' not in st.session_state:
        st.session_state.original_name = None

    # Create two-column layout for Steps 1 and 2
    col_step1, col_step2 = st.columns(2)

    # Step 1: Enhanced Search Section
    with col_step1:
        st.subheader("🔍 Step 1: Search for a Record")

        col_search, col_filter = st.columns([2, 1])

        with col_search:
            search_name = st.text_input(
                "Search by name:",
                placeholder="Type a name to search...",
                help="Smart search: supports partial matches, multiple words, and fuzzy matching (e.g., 'ana izq' finds 'ana izquierdo')",
                key="search_input"
            )

        with col_filter:
            show_all = st.checkbox("Show all records", value=False, key="show_all_checkbox")

        # Enhanced search with fuzzy matching and multi-word support
        if search_name and not show_all:
            # Use intelligent fuzzy search that handles partial matches, multiple words,
            # and similarity matching (e.g., "ana izq" finds "ana izquierdo")
            filtered_df = fuzzy_search_names(df, search_name)
        elif show_all:
            filtered_df = df
        else:
            filtered_df = df.head(0)  # Empty dataframe when no search criteria

    # Step 2: Select Record
    with col_step2:
        st.subheader("📋 Step 2: Select a Record")

        if not filtered_df.empty:
            # Create selectable options
            record_options = []
            for idx, row in filtered_df.iterrows():
                name = row.get('name', 'Unknown')
                date_connected = row.get('date_connected')
                if pd.notna(date_connected):
                    date_str = date_connected.strftime('%Y-%m-%d')
                else:
                    date_str = 'No date'

                # Show key info for selection
                display_text = f"{name} - Connected: {date_str}"
                record_options.append(display_text)

            # Add a "Clear selection" option
            record_options.insert(0, "--- Select a record ---")

            # Auto-select first record if search results exist
            default_index = 1 if len(record_options) > 1 else 0  # Index 1 is the first actual record

            selected_option = st.selectbox(
                "Choose a record to edit:",
                options=record_options,
                index=default_index,
                help="Select the record you want to edit",
                key="record_selector"
            )

            if selected_option != "--- Select a record ---":
                # Find the selected record
                selected_idx = record_options.index(selected_option) - 1  # -1 because we inserted at position 0
                selected_row = filtered_df.iloc[selected_idx]

                # Store in session state - handle NaN values
                answered_value = selected_row.get('ind_answered', 0)
                if pd.isna(answered_value):
                    answered_value = 0

                # Handle NaN values for job title and reachout type
                job_title_value = selected_row.get('job_title', '')
                if pd.isna(job_title_value):
                    job_title_value = ''

                reachout_type_value = selected_row.get('reach_out_type', '')
                if pd.isna(reachout_type_value):
                    reachout_type_value = ''

                st.session_state.selected_record = {
                    'name': selected_row.get('name', ''),
                    'date_connected': selected_row.get('date_connected', None),
                    'ind_answered': answered_value,
                    'url_chat': selected_row.get('url_chat', ''),
                    'company': selected_row.get('company', ''),
                    'job_title': job_title_value,
                    'location': selected_row.get('location', ''),
                    'date_contacted': selected_row.get('date_contacted', None),
                    'date_revocation': selected_row.get('date_revocation', None),
                    'date_discarded': selected_row.get('date_discarded', None),
                    'reach_out_type': reachout_type_value
                }
                st.session_state.original_name = selected_row.get('name', '')

                st.success(f"✅ Selected: {selected_option}")
            else:
                st.session_state.selected_record = None
                st.session_state.original_name = None
        else:
            if search_name:
                st.warning("No records found matching your search.")
            st.session_state.selected_record = None
            st.session_state.original_name = None

    # Step 3: Edit Form
    st.subheader("📝 Step 3: Edit Record Data")

    if st.session_state.selected_record:
        # Widget keys are namespaced by the selected record: an explicit key makes
        # Streamlit reuse the stored widget state and ignore `value=`, so a key that
        # did not vary per record would show the previously edited record's values.
        record_key = str(st.session_state.original_name or "")

        with st.form("edit_form"):
            # First row: Name and Answered (same line)
            col_name, col_answered = st.columns(2)

            with col_name:
                name = st.text_input(
                    "Name *",
                    value=st.session_state.selected_record.get('name', ''),
                    help="Full name of the LinkedIn profile",
                    key=f"dataentry_name_{record_key}"
                )

            with col_answered:
                # Handle NaN values and ensure valid index
                answered_value = st.session_state.selected_record.get('ind_answered', 0)
                if pd.isna(answered_value):
                    answered_value = 0
                answered_index = int(answered_value) if answered_value in [0, 1] else 0

                answered = st.selectbox(
                    "Answered",
                    options=[0, 1],
                    index=answered_index,
                    help="0 = Not answered, 1 = Answered",
                    key=f"dataentry_answered_{record_key}"
                )

            # Second row: All four date fields with clear checkboxes (8 columns total)
            col_contacted, col_clear_contacted, col_connected, col_clear_connected, col_revocation, col_clear_revocation, col_discarded, col_clear_discarded = st.columns(8)

            with col_contacted:
                day_contacted = st.date_input(
                    "Contacted",
                    value=st.session_state.selected_record.get('date_contacted') if pd.notna(st.session_state.selected_record.get('date_contacted')) else None,
                    help="Date when the contact was made",
                    key=f"dataentry_date_contacted_{record_key}"
                )

            with col_clear_contacted:
                clear_contacted = st.checkbox(
                    "Clear",
                    key="clear_contacted",
                    help="Clear the date contacted field"
                )

            with col_connected:
                date_connected = st.date_input(
                    "Connected",
                    value=st.session_state.selected_record.get('date_connected') if pd.notna(st.session_state.selected_record.get('date_connected')) else None,
                    help="Date when the connection was made",
                    key=f"dataentry_date_connected_{record_key}"
                )

            with col_clear_connected:
                clear_connected = st.checkbox(
                    "Clear",
                    key="clear_connected",
                    help="Clear the date connected field"
                )

            with col_revocation:
                revocation_date = st.date_input(
                    "Revocation",
                    value=st.session_state.selected_record.get('date_revocation') if pd.notna(st.session_state.selected_record.get('date_revocation')) else None,
                    help="Date when the contact was revoked",
                    key=f"dataentry_date_revocation_{record_key}"
                )

            with col_clear_revocation:
                clear_revocation = st.checkbox(
                    "Clear",
                    key="clear_revocation",
                    help="Clear the revocation date"
                )

            with col_discarded:
                date_discarded = st.date_input(
                    "Discarded",
                    value=st.session_state.selected_record.get('date_discarded') if pd.notna(st.session_state.selected_record.get('date_discarded')) else None,
                    help="Date when the profile was discarded",
                    key=f"dataentry_date_discarded_{record_key}"
                )

            with col_clear_discarded:
                clear_discarded = st.checkbox(
                    "Clear",
                    key="clear_discarded",
                    help="Clear the discarded date"
                )

            # Third row: LinkedIn Chat URL
            chat_url = st.text_input(
                "LinkedIn Chat URL",
                value=st.session_state.selected_record.get('url_chat', ''),
                help="URL to the LinkedIn chat/messaging thread",
                key=f"dataentry_url_chat_{record_key}"
            )

            # Fourth row: Company and Reachout Type (same line)
            col_company, col_reachout = st.columns(2)

            with col_company:
                company = st.text_input(
                    "Company",
                    value=st.session_state.selected_record.get('company', ''),
                    help="Company name",
                    key=f"dataentry_company_{record_key}"
                )

            with col_reachout:
                # Reachout Type dropdown
                # Get distinct reach out types from the full dataframe
                reachout_types = ['']  # Start with empty option
                if 'reach_out_type' in df.columns:
                    distinct_types = df['reach_out_type'].dropna().unique().tolist()
                    reachout_types.extend(sorted(distinct_types))

                # Ensure current record's reachout type is in the options list
                current_reachout_type = st.session_state.selected_record.get('reach_out_type', '')
                if current_reachout_type and current_reachout_type not in reachout_types:
                    reachout_types.append(current_reachout_type)
                    reachout_types.sort()  # Keep sorted after adding

                # Find the index of the current value
                try:
                    reachout_index = reachout_types.index(current_reachout_type) if current_reachout_type else 0
                except ValueError:
                    reachout_index = 0

                reach_out_type = st.selectbox(
                    "Reachout Type",
                    options=reachout_types,
                    index=reachout_index,
                    help="Type of reachout made to this profile",
                    key=f"dataentry_reach_out_type_{record_key}"
                )

            # Fifth row: Job Title and Location (same line)
            col_job_title, col_location = st.columns(2)

            with col_job_title:
                job_title = st.text_input(
                    "Job Title",
                    value=st.session_state.selected_record.get('job_title', ''),
                    help="Job title/position",
                    key=f"dataentry_job_title_{record_key}"
                )

            with col_location:
                location = st.text_input(
                    "Location",
                    value=st.session_state.selected_record.get('location', ''),
                    help="Location/city",
                    key=f"dataentry_location_{record_key}"
                )

            submitted = st.form_submit_button("💾 Update Record", key="dataentry_submit")

            if submitted:
                if not name.strip():
                    st.error("❌ Name is required!")
                    return

                # Prepare record data
                # Handle clear checkboxes - if checked, set date to None
                final_date_connected = None if clear_connected else (pd.Timestamp(date_connected) if date_connected else None)
                final_day_contacted = None if clear_contacted else (pd.Timestamp(day_contacted) if day_contacted else None)
                final_revocation_date = None if clear_revocation else (pd.Timestamp(revocation_date) if revocation_date else None)
                final_date_discarded = None if clear_discarded else (pd.Timestamp(date_discarded) if date_discarded else None)

                # Reset Logic: If contact date is updated (and not just cleared), reset revocation date
                original_day = st.session_state.selected_record.get('date_contacted')
                # Check if day has changed to a new valid date (re-contact logic)
                if final_day_contacted is not None and (pd.isna(original_day) or final_day_contacted != original_day):
                     final_revocation_date = None
                     if revocation_date is not None or clear_revocation:
                         st.info("ℹ️ Revocation date automatically cleared because a new contact date was set.")

                record_data = {
                    'name': name.strip(),
                    'date_connected': final_date_connected,
                    'ind_answered': answered,
                    'url_chat': chat_url.strip() if chat_url else None,
                    'company': company.strip() if company else None,
                    'job_title': job_title.strip() if job_title else None,
                    'location': location.strip() if location else None,
                    'date_contacted': final_day_contacted,
                    'date_revocation': final_revocation_date,
                    'date_discarded': final_date_discarded,
                    'reach_out_type': reach_out_type.strip() if reach_out_type else None
                }

                # Update existing record
                updated_df, action = update_existing_record(df, st.session_state.original_name, record_data)

                if action == "not_found":
                    st.error("❌ Original record not found. It may have been deleted.")
                    return

                # Save to Excel
                if save_to_excel(updated_df, data_path):
                    st.success("✅ Record updated successfully!")
                    
                    # Log history with original data for initial record
                    log_history(record_data, action="update", original_data=st.session_state.selected_record)

                    # Clear selection and rerun
                    st.session_state.selected_record = None
                    st.session_state.original_name = None
                    st.rerun()
                else:
                    st.error("❌ Failed to save changes!")
    else:
        st.info("👆 Please select a record from Step 2 to edit its data.")

    # Step 4: Manual formatting button
    st.divider()
    st.subheader("🎨 Excel Formatting")
    
    col_format_info, col_format_btn = st.columns([3, 1])
    
    with col_format_info:
        st.markdown("""
        Apply Excel formatting to convert URLs to hyperlinks and format columns.
        **Note:** This process takes ~60 seconds. Only click when you're done editing.
        """)
    
    with col_format_btn:
        if st.button("🎨 Apply Formatting", type="primary", width="stretch", key="dataentry_apply_formatting"):
            json_path = Path(__file__).parent / "excel_format_spec.json"
            if not json_path.exists():
                st.error("❌ Format specification file not found!")
            else:
                with st.spinner("Applying Excel formatting... This may take up to 60 seconds."):
                    try:
                        if apply_excel_formatting(data_path, str(json_path)):
                            st.success("✅ Excel formatting applied successfully!")
                        else:
                            st.error("❌ Failed to apply Excel formatting. Check the logs for details.")
                    except Exception as e:
                        st.error(f"❌ Error applying Excel formatting: {e}")


if __name__ == "__main__":
    main()