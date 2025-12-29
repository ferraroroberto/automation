import json
import os
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st

# Import Excel formatting functions
from excel_format_manager import convert_url_columns_to_hyperlinks, apply_format_from_json

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
        date_columns = ['day', 'date connected']
        for col in date_columns:
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

def main():
    """Data entry page for editing LinkedIn profiles."""

    st.title("✏️ LinkedIn Profile Data Editor")

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

    # Display current record count
    st.info(f"📊 Current records: {len(df)}")

    # Initialize session state for selected record
    if 'selected_record' not in st.session_state:
        st.session_state.selected_record = None
    if 'original_name' not in st.session_state:
        st.session_state.original_name = None

    # Step 1: Search Section
    st.subheader("🔍 Step 1: Search for a Record")

    col_search, col_filter = st.columns([2, 1])

    with col_search:
        search_name = st.text_input(
            "Search by name:",
            placeholder="Type a name to search...",
            help="Search for existing profiles by name",
            key="search_input"
        )

    with col_filter:
        show_all = st.checkbox("Show all records", value=False, key="show_all_checkbox")

    # Filter data based on search
    if search_name and not show_all:
        # Case-insensitive search
        mask = df['name'].str.lower().str.contains(search_name.lower(), na=False) if 'name' in df.columns else pd.Series([False] * len(df))
        filtered_df = df[mask]
    elif show_all:
        filtered_df = df
    else:
        filtered_df = df.head(0)  # Empty dataframe

    # Step 2: Select Record
    if not filtered_df.empty:
        st.subheader("📋 Step 2: Select a Record")

        # Create selectable options
        record_options = []
        for idx, row in filtered_df.iterrows():
            name = row.get('name', 'Unknown')
            date_connected = row.get('date connected')
            if pd.notna(date_connected):
                date_str = date_connected.strftime('%Y-%m-%d')
            else:
                date_str = 'No date'

            # Show key info for selection
            display_text = f"{name} - Connected: {date_str}"
            record_options.append(display_text)

        # Add a "Clear selection" option
        record_options.insert(0, "--- Select a record ---")

        selected_option = st.selectbox(
            "Choose a record to edit:",
            options=record_options,
            index=0,
            help="Select the record you want to edit",
            key="record_selector"
        )

        if selected_option != "--- Select a record ---":
            # Find the selected record
            selected_idx = record_options.index(selected_option) - 1  # -1 because we inserted at position 0
            selected_row = filtered_df.iloc[selected_idx]

            # Store in session state - handle NaN values
            answered_value = selected_row.get('answered', 0)
            if pd.isna(answered_value):
                answered_value = 0

            st.session_state.selected_record = {
                'name': selected_row.get('name', ''),
                'date_connected': selected_row.get('date connected', None),
                'answered': answered_value,
                'chat_url': selected_row.get('chat_url', '')
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

    st.markdown("---")

    # Step 3: Edit Form
    st.subheader("📝 Step 3: Edit Record Data")

    if st.session_state.selected_record:
        with st.form("edit_form"):
            col1, col2 = st.columns(2)

            with col1:
                name = st.text_input(
                    "Name *",
                    value=st.session_state.selected_record.get('name', ''),
                    help="Full name of the LinkedIn profile"
                )

                date_connected = st.date_input(
                    "Date Connected",
                    value=st.session_state.selected_record.get('date_connected') if pd.notna(st.session_state.selected_record.get('date_connected')) else None,
                    help="Date when the connection was made"
                )

            with col2:
                # Handle NaN values and ensure valid index
                answered_value = st.session_state.selected_record.get('answered', 0)
                if pd.isna(answered_value):
                    answered_value = 0
                answered_index = int(answered_value) if answered_value in [0, 1] else 0

                answered = st.selectbox(
                    "Answered",
                    options=[0, 1],
                    index=answered_index,
                    help="0 = Not answered, 1 = Answered"
                )

                chat_url = st.text_input(
                    "LinkedIn Chat URL",
                    value=st.session_state.selected_record.get('chat_url', ''),
                    help="URL to the LinkedIn chat/messaging thread"
                )

            submitted = st.form_submit_button("💾 Update Record")

            if submitted:
                if not name.strip():
                    st.error("❌ Name is required!")
                    return

                # Prepare record data
                record_data = {
                    'name': name.strip(),
                    'date connected': pd.Timestamp(date_connected) if date_connected else None,
                    'answered': answered,
                    'chat_url': chat_url.strip() if chat_url else None
                }

                # Update existing record
                updated_df, action = update_existing_record(df, st.session_state.original_name, record_data)

                if action == "not_found":
                    st.error("❌ Original record not found. It may have been deleted.")
                    return

                # Save to Excel
                if save_to_excel(updated_df, data_path):
                    st.success("✅ Record updated successfully!")

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
                    st.session_state.selected_record = None
                    st.session_state.original_name = None
                    st.rerun()
                else:
                    st.error("❌ Failed to save changes!")
    else:
        st.info("👆 Please select a record from Step 2 to edit its data.")

    # Recent Records Section
    st.markdown("---")
    st.subheader("📋 Recent Records")

    # Show last 5 records
    if not df.empty:
        recent_df = df.tail(5).copy()

        # Format dates for display
        if 'date connected' in recent_df.columns:
            recent_df['date connected'] = recent_df['date connected'].dt.strftime('%Y-%m-%d')

        st.dataframe(recent_df)
    else:
        st.info("No records yet.")

if __name__ == "__main__":
    main()