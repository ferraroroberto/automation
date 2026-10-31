
import pandas as pd
import streamlit as st
import plotly.express as px

from history_manager import log_history
from loaders import load_config
from _lib import (
    fuzzy_search_profiles,
    save_to_excel,
    update_existing_record,
    render_excel_formatting_section,
    generate_color_gradient,
)


def _select_view_df(df_filtered, view_mode):
    """Return the rows of the chosen view (uncontacted, discarded or revoked, with the aging filter)."""
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
    return uncontacted_df


def _show_empty_message(view_mode):
    """Say why the chosen view has no rows."""
    if view_mode == "Uncontacted Profiles":
        st.info("🎉 All profiles have been contacted! No reachout targets remaining.")
    elif view_mode == "Discarded Contacts":
        st.info("No discarded contacts found.")
    else:
        st.info("No revoked contacts found matching criteria.")


def _render_company_chart(uncontacted_df, chart_title):
    """Show the stacked bar chart of the view's rows by company and search type."""
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


def _render_profile_list(filtered_uncontacted_df):
    """List the profiles as buttons; clicking one selects it for editing."""
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


def _discard_profile(df_all, data_path):
    """Stamp the selected profile as discarded today and save it."""
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


def _as_date_or_none(value):
    """Normalise a stored date (Timestamp, string, NaT or empty) to a date for st.date_input."""
    if pd.notna(value) and value != '':
        try:
            # Convert to datetime.date if it's a pandas Timestamp
            if hasattr(value, 'date'):
                value = value.date()
            elif isinstance(value, str):
                value = pd.to_datetime(value).date()
        except (ValueError, TypeError):
            value = None
    else:
        value = None
    return value


def _reachout_type_options(df_all):
    """Selectbox options: an empty choice plus the distinct reachout types in the full data."""
    reachout_types = ['']  # Start with empty option
    if 'reach_out_type' in df_all.columns:
        distinct_types = df_all['reach_out_type'].dropna().unique().tolist()
        reachout_types.extend(sorted(distinct_types))
    return reachout_types


def _render_reachout_fields(record, record_key, df_all):
    """Render the reachout edit-form fields for one profile; returns their values."""
    # Editable fields
    name = st.text_input(
        "Name *",
        value=record.get('name', ''),
        help="Full name of the LinkedIn profile",
        key=f"reachout_name_{record_key}"
    )

    date_contacted_value = _as_date_or_none(record.get('date_contacted', ''))

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

    date_discarded_value = _as_date_or_none(record.get('date_discarded', ''))

    date_discarded = st.date_input(
        "Date Discarded",
        value=date_discarded_value,
        help="Date when this profile was discarded (leave empty if not discarded)",
        key=f"reachout_date_discarded_{record_key}"
    )

    reachout_types = _reachout_type_options(df_all)

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

    return {
        'name': name,
        'day': day,
        'date_discarded': date_discarded,
        'reach_out_type': reach_out_type,
        'search_type': search_type,
        'job_title': job_title,
        'follows_from': follows_from,
        'company': company,
        'location': location,
        'url': url,
    }


def _save_reachout_edit(df_all, data_path, fields):
    """Validate the submitted form and save it; False when the page should stop rendering."""
    (name, day, date_discarded, reach_out_type, search_type,
     job_title, follows_from, company, location, url) = (
        fields[k] for k in (
            'name', 'day', 'date_discarded', 'reach_out_type', 'search_type',
            'job_title', 'follows_from', 'company', 'location', 'url'))

    if not name.strip():
        st.error("❌ Name is required!")
        return False

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
        return False

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
    return True


def _render_reachout_form(df_all, data_path):
    """Render the edit form for the selected profile; False when the page should stop rendering."""
    record = st.session_state.reachout_selected_record
    # Widget keys are namespaced by the selected profile: an explicit key makes
    # Streamlit reuse the stored widget state and ignore `value=`, so a key that
    # did not vary per profile would show the previously edited profile's values.
    record_key = str(st.session_state.reachout_original_name or "")

    with st.form("reachout_edit_form"):
        fields = _render_reachout_fields(record, record_key, df_all)

        submitted = st.form_submit_button("💾 Save Changes", key="reachout_submit")

        if submitted:
            return _save_reachout_edit(df_all, data_path, fields)
    return True


def _delete_reachout_profile(df_all, data_path, name_to_delete, pending_key):
    """Remove the named profile from the workbook, log it to history and reset the edit panel."""
    try:
        mask = df_all['name'].str.lower() == name_to_delete.lower() if 'name' in df_all.columns else pd.Series([False] * len(df_all))
        if not mask.any():
            st.error("❌ Profile not found for deletion.")
            return

        # Get record data for history before deleting
        record_to_delete = df_all[mask].iloc[0].to_dict()
        df_all = df_all[~mask].copy()

        if not save_to_excel(df_all, data_path):
            st.error("❌ Failed to save changes after deletion!")
            return

        st.success(f"✅ Profile '{name_to_delete}' deleted successfully!")
        log_history(record_to_delete, action="delete")

        st.session_state.reachout_selected_record = None
        st.session_state.reachout_original_name = None
        st.session_state.reachout_editing = False
        st.session_state.pop(pending_key, None)
        st.rerun()
    except Exception as e:
        st.error(f"❌ Error deleting profile: {str(e)}")


def _render_delete_section(df_all, data_path):
    """Show the delete button with its two-step confirmation.

    The confirmation is rendered from a session-state flag, not from inside the
    first button's branch: a click reruns the script, and a nested button is
    only evaluated while its parent button is still True.
    """
    name = st.session_state.reachout_original_name
    pending_key = f"confirm_delete_{name}"

    st.markdown("---")
    col_delete, col_spacer = st.columns([1, 3])
    with col_delete:
        if st.button("🗑️ Delete Profile", key="delete_profile", type="secondary",
                   help="Permanently delete this profile from the database"):
            st.session_state[pending_key] = True

        if st.session_state.get(pending_key):
            st.warning(f"⚠️ Are you sure you want to delete '{name}'? This action cannot be undone!")
            if st.button("✅ Yes, Delete", key=f"confirm_yes_{name}"):
                _delete_reachout_profile(df_all, data_path, name, pending_key)
            if st.button("❌ Cancel", key=f"confirm_no_{name}"):
                del st.session_state[pending_key]
                st.rerun()


def _render_edit_panel(df_all, data_path):
    """Render the edit column (header, discard, form, delete); False when the page should stop rendering."""
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
                _discard_profile(df_all, data_path)

    if st.session_state.reachout_editing and st.session_state.reachout_selected_record:
        if not _render_reachout_form(df_all, data_path):
            return False

        _render_delete_section(df_all, data_path)
    else:
        st.info("👆 Click on a profile from the list to edit its information.")
    return True


def _chart_title(view_mode):
    """Heading noun for the chosen view."""
    if view_mode == "Uncontacted Profiles":
        chart_title = "Uncontacted Profiles"
    elif view_mode == "Discarded Contacts":
        chart_title = "Discarded Contacts"
    else:
        chart_title = "Revoked Contacts"
    return chart_title


def _init_selection_state():
    """Create the session-state slots that hold the selected profile."""
    if 'reachout_selected_record' not in st.session_state:
        st.session_state.reachout_selected_record = None
    if 'reachout_original_name' not in st.session_state:
        st.session_state.reachout_original_name = None
    if 'reachout_editing' not in st.session_state:
        st.session_state.reachout_editing = False


def _search_view_rows(uncontacted_df, view_mode):
    """Render the search box and the show-all toggle; returns the search text, the show-all state and the matching rows."""
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
    return search_uncontacted, show_all_uncontacted, filtered_uncontacted_df


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

    uncontacted_df = _select_view_df(df_filtered, view_mode)

    if uncontacted_df.empty:
        _show_empty_message(view_mode)
        return

    # Add horizontal bar chart showing uncontacted people by company at the top
    chart_title = _chart_title(view_mode)
    st.subheader(f"📈 {chart_title} by Company")

    _render_company_chart(uncontacted_df, chart_title)

    st.markdown("---")

    search_uncontacted, show_all_uncontacted, filtered_uncontacted_df = _search_view_rows(uncontacted_df, view_mode)

    # Display filtered count
    if search_uncontacted or show_all_uncontacted:
        st.subheader(f"📋 {chart_title} ({len(filtered_uncontacted_df)} filtered from {len(uncontacted_df)} total)")
    else:
        st.subheader(f"📋 {chart_title} ({len(uncontacted_df)} total)")

    _init_selection_state()

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

        _render_profile_list(filtered_uncontacted_df)

    with col_edit:
        if not _render_edit_panel(df_all, data_path):
            return

    # Show summary table at the bottom
    st.markdown("---")
    st.subheader("📊 Reachout Summary")

    if not filtered_uncontacted_df.empty:
        # Display summary table with the requested columns
        available_cols = [col for col in filtered_uncontacted_df.columns]

        if available_cols:
            summary_df = filtered_uncontacted_df[available_cols].copy()
            st.dataframe(summary_df, width='stretch')
        else:
            st.warning("Required columns not found in data.")

    # Manual formatting button at the bottom
    render_excel_formatting_section(data_path, "reachout_format_btn")


if __name__ == "__main__":
    main()