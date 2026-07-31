from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st
import plotly.express as px

from history_manager import log_history
from loaders import load_config, load_excel_data
from _lib import (
    fuzzy_search_profiles,
    save_to_excel,
    apply_excel_formatting,
    generate_color_gradient,
)

def update_existing_record(df, original_name, record_data):
    """Update an existing record by original name."""
    # Find the record by original name (case-insensitive)
    existing_mask = df['name'].str.lower() == original_name.lower() if 'name' in df.columns else pd.Series([False] * len(df))

    if existing_mask.any():
        # Update existing record
        idx = df[existing_mask].index[0]
        # Define date columns that need special handling
        date_columns = ['date_contacted', 'date_connected', 'date_revocation', 'date_discarded']
        
        for key, value in record_data.items():
            if key in df.columns:
                # Handle date columns: ensure proper dtype conversion
                if key in date_columns:
                    # Convert value to datetime if it's not already NaT/None
                    if pd.notna(value) and value != '':
                        # Ensure the column is datetime dtype
                        if not pd.api.types.is_datetime64_any_dtype(df[key]):
                            df[key] = pd.to_datetime(df[key], errors='coerce')
                        # Convert value to Timestamp
                        if isinstance(value, str):
                            value = pd.to_datetime(value, errors='coerce')
                        elif not isinstance(value, pd.Timestamp):
                            value = pd.Timestamp(value) if value else pd.NaT
                    else:
                        value = pd.NaT
                
                df.at[idx, key] = value
        return df, "updated"
    else:
        return df, "not_found"

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
    view_mode = st.radio("View Mode", ["Uncontacted Profiles", "Revoked Contacts", "Discarded Contacts"], horizontal=True, key="reachout_view_mode")

    if view_mode == "Uncontacted Profiles":
        # Filter for profiles that are not contacted AND not discarded
        if 'date_contacted' in df_filtered.columns and 'date_discarded' in df_filtered.columns:
            uncontacted_df = df_filtered[df_filtered['date_contacted'].isna() & df_filtered['date_discarded'].isna()].copy()
        elif 'date_contacted' in df_filtered.columns:
            uncontacted_df = df_filtered[df_filtered['date_contacted'].isna()].copy()
        else:
            uncontacted_df = df_filtered.copy()
        
        # Filter out records that are actually revoked (have a revocation date)
        if 'date_revocation' in uncontacted_df.columns:
             # Just to be safe, though they shouldn't have a 'date_contacted' anyway
             pass
             
    elif view_mode == "Discarded Contacts":
        # Filter for profiles that have been discarded (no timeframe filter)
        if 'date_discarded' in df_filtered.columns:
            uncontacted_df = df_filtered[df_filtered['date_discarded'].notna()].copy()
        else:
            uncontacted_df = pd.DataFrame()
            st.warning("Discarded date column missing from data.")
            
    else: # Revoked Contacts
        if 'date_revocation' in df_filtered.columns and 'date_contacted' in df_filtered.columns:
            # Filter where date_revocation is present AND date_revocation > date_contacted
            # Note: We must handle cases where 'date_contacted' might be NaT or populated.
            # Logic: A contact is "revoked" if it HAS been contacted (date_contacted is not null)
            # AND it has a revocation date that is AFTER the contact date.

            mask = (df_filtered['date_contacted'].notna()) & (df_filtered['date_revocation'].notna()) & (df_filtered['date_revocation'] > df_filtered['date_contacted'])
            uncontacted_df = df_filtered[mask].copy()
            
            # Aging Filters for Revoked Contacts (not for Discarded Contacts)
            st.write("🕒 **Aging Filter** (Time since revocation)")
            col_age1, col_age2, col_age3, col_age4 = st.columns(4)
            
            age_filter = st.radio(
                "Show profiles revoked at least:",
                ["All Revoked", "4 Weeks Ago", "8 Weeks Ago", "12 Weeks Ago"],
                horizontal=True,
                label_visibility="collapsed",
                key="reachout_age_filter"
            )
            
            if age_filter != "All Revoked":
                now = pd.Timestamp.now()
                weeks = 0
                if age_filter == "4 Weeks Ago": weeks = 4
                elif age_filter == "8 Weeks Ago": weeks = 8
                elif age_filter == "12 Weeks Ago": weeks = 12
                
                cutoff_date = now - pd.Timedelta(weeks=weeks)
                uncontacted_df = uncontacted_df[uncontacted_df['date_revocation'] <= cutoff_date]
                
        else:
            uncontacted_df = pd.DataFrame()
            st.warning("Revocation date column missing from data.")

    if uncontacted_df.empty:
        if view_mode == "Uncontacted Profiles":
            st.info("🎉 All profiles have been contacted! No reachout targets remaining.")
        elif view_mode == "Discarded Contacts":
            st.info("No discarded contacts found.")
        else:
            st.info("No revoked contacts found matching criteria.")
        return

    # Add horizontal bar chart showing uncontacted people by company at the top
    if view_mode == "Uncontacted Profiles":
        chart_title = "Uncontacted Profiles"
    elif view_mode == "Discarded Contacts":
        chart_title = "Discarded Contacts"
    else:
        chart_title = "Revoked Contacts"
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
                    'date_contacted': row.get('date_contacted', ''),
                    'job_title': job_title,
                    'follows_from': row.get('follows_from', ''),
                    'company': company,
                    'location': location,
                    'search_type': row.get('search_type', ''),
                    'url_profile': row.get('url_profile', ''),
                    'reach_out_type': row.get('reach_out_type', ''),
                    'date_revocation': row.get('date_revocation', None),
                    'date_discarded': row.get('date_discarded', None)
                }
                st.session_state.reachout_original_name = name
                st.session_state.reachout_editing = True
                st.rerun()

    with col_edit:
        # Header with Edit Profile, Open Profile link, and Discard button
        col_header, col_link, col_discard = st.columns([2, 1, 1])
        with col_header:
            st.subheader("✏️ Edit Profile")
        with col_link:
            if st.session_state.reachout_editing and st.session_state.reachout_selected_record:
                profile_url = st.session_state.reachout_selected_record.get('url_profile', '')
                if profile_url:
                    st.markdown(f'<a href="{profile_url}" target="_blank" style="text-decoration: none;"><button style="background-color: #0077b5; color: white; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer;">🔗 Open Profile</button></a>', unsafe_allow_html=True)
        with col_discard:
            if st.session_state.reachout_editing and st.session_state.reachout_selected_record:
                if st.button("🗑️ Discard", key="discard_profile_btn", help="Mark this profile as discarded", width="stretch"):
                    # Update date_discarded to today
                    record_data = st.session_state.reachout_selected_record.copy()
                    record_data['date_discarded'] = pd.Timestamp.now()
                    
                    # Update existing record in the dataframe
                    updated_df, action = update_existing_record(df_all, st.session_state.reachout_original_name, record_data)
                    
                    if action == "not_found":
                        st.error("❌ Original record not found. It may have been deleted.")
                    else:
                        # Save to Excel
                        if save_to_excel(updated_df, data_path):
                            st.success("✅ Profile discarded successfully!")
                            
                            # Log history with action "discard"
                            log_history(record_data, action="discard", original_data=st.session_state.reachout_selected_record)
                            
                            # Clear selection and rerun
                            st.session_state.reachout_selected_record = None
                            st.session_state.reachout_original_name = None
                            st.session_state.reachout_editing = False
                            st.rerun()
                        else:
                            st.error("❌ Failed to save changes!")

        if st.session_state.reachout_editing and st.session_state.reachout_selected_record:
            record = st.session_state.reachout_selected_record
            # Widget keys are namespaced by the selected profile: an explicit key makes
            # Streamlit reuse the stored widget state and ignore `value=`, so a key that
            # did not vary per profile would show the previously edited profile's values.
            record_key = str(st.session_state.reachout_original_name or "")

            with st.form("reachout_edit_form"):


                # Editable fields
                name = st.text_input(
                    "Name *",
                    value=record.get('name', ''),
                    help="Full name of the LinkedIn profile",
                    key=f"reachout_name_{record_key}"
                )

                # Handle date_contacted field - convert to date if it's not None/NaT
                date_contacted_value = record.get('date_contacted', '')
                if pd.notna(date_contacted_value) and date_contacted_value != '':
                    try:
                        # Convert to datetime.date if it's a pandas Timestamp
                        if hasattr(date_contacted_value, 'date'):
                            date_contacted_value = date_contacted_value.date()
                        elif isinstance(date_contacted_value, str):
                            date_contacted_value = pd.to_datetime(date_contacted_value).date()
                    except (ValueError, TypeError):
                        date_contacted_value = None
                else:
                    date_contacted_value = None

                day = st.date_input(
                    "Day Contacted",
                    value=date_contacted_value,
                    help="Date when this profile was contacted (leave empty if not contacted yet)",
                    key=f"reachout_day_contacted_{record_key}"
                )

                # Show revocation info if present
                revocation_val = record.get('date_revocation')
                if pd.notna(revocation_val):
                    rev_str = revocation_val.strftime('%Y-%m-%d') if hasattr(revocation_val, 'strftime') else str(revocation_val)
                    st.warning(f"⚠️ This contact was revoked on {rev_str}. Updating the 'Day Contacted' will clear the revocation date.")

                # Handle date_discarded field - convert to date if it's not None/NaT
                date_discarded_value = record.get('date_discarded', '')
                if pd.notna(date_discarded_value) and date_discarded_value != '':
                    try:
                        # Convert to datetime.date if it's a pandas Timestamp
                        if hasattr(date_discarded_value, 'date'):
                            date_discarded_value = date_discarded_value.date()
                        elif isinstance(date_discarded_value, str):
                            date_discarded_value = pd.to_datetime(date_discarded_value).date()
                    except (ValueError, TypeError):
                        date_discarded_value = None
                else:
                    date_discarded_value = None

                date_discarded = st.date_input(
                    "Date Discarded",
                    value=date_discarded_value,
                    help="Date when this profile was discarded (leave empty if not discarded)",
                    key=f"reachout_date_discarded_{record_key}"
                )

                # Get distinct reach out types from the full dataframe
                reachout_types = ['']  # Start with empty option
                if 'reach_out_type' in df_all.columns:
                    distinct_types = df_all['reach_out_type'].dropna().unique().tolist()
                    reachout_types.extend(sorted(distinct_types))

                reach_out_type = st.selectbox(
                    "Reachout Type",
                    options=reachout_types,
                    index=reachout_types.index(record.get('reach_out_type', '')) if record.get('reach_out_type', '') in reachout_types else 0,
                    help="Type of reachout made to this profile",
                    key=f"reachout_reach_out_type_{record_key}"
                )

                search_type = st.text_input(
                    "Search Type",
                    value=record.get('search_type', ''),
                    help="How this profile was found (e.g., keyword search, mutual connections, etc.)",
                    key=f"reachout_search_type_{record_key}"
                )

                job_title = st.text_input(
                    "Job Title",
                    value=record.get('job_title', ''),
                    help="Current job position",
                    key=f"reachout_job_title_{record_key}"
                )

                follows_from = st.text_input(
                    "Follows From",
                    value=record.get('follows_from', ''),
                    help="How you connected with this person",
                    key=f"reachout_follows_from_{record_key}"
                )

                company = st.text_input(
                    "Company",
                    value=record.get('company', ''),
                    help="Company name",
                    key=f"reachout_company_{record_key}"
                )

                location = st.text_input(
                    "Location",
                    value=record.get('location', ''),
                    help="Location/city",
                    key=f"reachout_location_{record_key}"
                )

                url = st.text_input(
                    "Profile URL",
                    value=record.get('url_profile', ''),
                    help="LinkedIn profile URL",
                    key=f"reachout_url_profile_{record_key}"
                )

                submitted = st.form_submit_button("💾 Save Changes", key="reachout_submit")

                if submitted:
                    if not name.strip():
                        st.error("❌ Name is required!")
                        return

                    # Prepare record data
                    final_day = pd.Timestamp(day) if day else pd.NaT
                    final_date_discarded = pd.Timestamp(date_discarded) if date_discarded else pd.NaT
                    
                    # Reset Logic: If contact date is updated (and not just cleared), reset revocation date
                    original_day = st.session_state.reachout_selected_record.get('date_contacted')
                    # Handle NaT comparison
                    original_day_ts = pd.Timestamp(original_day) if pd.notna(original_day) else pd.NaT

                    # Check if day has changed to a new valid date (re-contact logic)
                    revocation_update = st.session_state.reachout_selected_record.get('date_revocation')
                    if day and (pd.isna(original_day_ts) or final_day != original_day_ts):
                         revocation_update = None
                         # We don't need to show a message here as the save success message is enough,
                         # but we ensure the dict sends None for date_revocation

                    record_data = {
                        'name': name.strip(),
                        'date_contacted': final_day,
                        'job_title': job_title.strip(),
                        'follows_from': follows_from.strip(),
                        'company': company.strip(),
                        'location': location.strip(),
                        'search_type': search_type.strip(),
                        'url_profile': url.strip(),
                        'reach_out_type': reach_out_type.strip() if reach_out_type else '',
                        'date_revocation': revocation_update,
                        'date_discarded': final_date_discarded
                    }

                    # Update existing record
                    updated_df, action = update_existing_record(df_all, st.session_state.reachout_original_name, record_data)

                    if action == "not_found":
                        st.error("❌ Original record not found. It may have been deleted.")
                        return

                    # Save to Excel
                    if save_to_excel(updated_df, data_path):
                        st.success("✅ Profile updated successfully!")
                        
                        # Log history with original data for initial record
                        log_history(record_data, action="update", original_data=st.session_state.reachout_selected_record)

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

    # Manual formatting button at the bottom
    st.divider()
    st.subheader("🎨 Excel Formatting")
    
    col_format_info, col_format_btn = st.columns([3, 1])
    
    with col_format_info:
        st.markdown("""
        Apply Excel formatting to convert URLs to hyperlinks and format columns.
        **Note:** This process takes ~60 seconds. Only click when you're done editing.
        """)
    
    with col_format_btn:
        if st.button("🎨 Apply Formatting", type="primary", width="stretch", key="reachout_format_btn"):
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