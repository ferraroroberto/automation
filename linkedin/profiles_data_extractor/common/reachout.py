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
    uncontacted_df = df_filtered[df_filtered['day'].isna()].copy() if 'day' in df_filtered.columns else df_filtered.copy()

    if uncontacted_df.empty:
        st.info("🎉 All profiles have been contacted! No reachout targets remaining.")
        return

    # Display filtered count
    st.subheader(f"📋 Uncontacted Profiles ({len(uncontacted_df)} total)")

    # Initialize session state for selected record
    if 'reachout_selected_record' not in st.session_state:
        st.session_state.reachout_selected_record = None
    if 'reachout_original_name' not in st.session_state:
        st.session_state.reachout_original_name = None
    if 'reachout_editing' not in st.session_state:
        st.session_state.reachout_editing = False

    if uncontacted_df.empty:
        st.info("No profiles match the selected filters.")
        return

    # Create two columns layout
    col_list, col_edit = st.columns([1, 1])

    with col_list:
        st.subheader("👥 Select Profile")

        # Create selectable list
        for idx, row in uncontacted_df.iterrows():
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
                    'reach out type': row.get('reach out type', '')
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
                    "Day",
                    value=day_value,
                    help="Date when this profile was contacted (leave empty if not contacted yet)"
                )

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
                    record_data = {
                        'name': name.strip(),
                        'day': pd.Timestamp(day) if day else pd.NaT,
                        'job_title': job_title.strip(),
                        'follows_from': follows_from.strip(),
                        'company': company.strip(),
                        'location': location.strip(),
                        'search_type': search_type.strip(),
                        'url': url.strip(),
                        'reach out type': reach_out_type.strip() if reach_out_type else ''
                    }

                    # Update existing record
                    updated_df, action = update_existing_record(df_all, st.session_state.reachout_original_name, record_data)

                    if action == "not_found":
                        st.error("❌ Original record not found. It may have been deleted.")
                        return

                    # Save to Excel
                    if save_to_excel(updated_df, data_path):
                        st.success("✅ Profile updated successfully!")

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
        else:
            st.info("👆 Click on a profile from the list to edit its information.")

    # Show summary table at the bottom
    st.markdown("---")
    st.subheader("📊 Reachout Summary")

    if not uncontacted_df.empty:
        # Display summary table with the requested columns
        display_cols = ['name', 'job_title', 'follows_from', 'company', 'location']
        available_cols = [col for col in display_cols if col in uncontacted_df.columns]

        if available_cols:
            summary_df = uncontacted_df[available_cols].copy()
            st.dataframe(summary_df, width='stretch')
        else:
            st.warning("Required columns not found in data.")

if __name__ == "__main__":
    main()