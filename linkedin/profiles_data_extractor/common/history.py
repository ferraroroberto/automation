import os
import pandas as pd
import streamlit as st
import openpyxl

# Import history manager functions
from history_manager import get_history_file_path
from loaders import load_config
from _lib import fuzzy_search_names

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


def delete_history_records(indices_to_delete):
    """
    Delete specific records from the history Excel file by their original indices.
    
    Args:
        indices_to_delete: List of original DataFrame indices to delete
        
    Returns:
        Tuple of (success: bool, message: str)
    """
    history_path = get_history_file_path()
    if not history_path or not os.path.exists(history_path):
        return False, "History file not found"
    
    try:
        # Load the raw history file (unsorted)
        df = pd.read_excel(history_path, engine='openpyxl')
        
        # Ensure we have valid indices
        valid_indices = [idx for idx in indices_to_delete if idx in df.index]
        
        if not valid_indices:
            return False, "No valid records to delete"
        
        # Drop the selected rows
        df = df.drop(valid_indices)
        
        # Reset index after deletion
        df = df.reset_index(drop=True)
        
        # Save back to Excel
        with pd.ExcelWriter(history_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='History', index=False)
        
        return True, f"Successfully deleted {len(valid_indices)} record(s)"
    
    except Exception as e:
        return False, f"Error deleting records: {e}"

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

            # Show recent activity with multi-select for deletion
            st.subheader("🕐 Recent Activity")
            
            # Get recent records with their original indices preserved
            recent_df = history_df.head(50)[['timestamp', 'action', 'name', 'company']].copy()
            
            # Store original indices for deletion
            original_indices = recent_df.index.tolist()
            
            # Format timestamps for display
            display_recent_df = recent_df.copy()
            valid_timestamps = pd.to_datetime(display_recent_df['timestamp'], errors='coerce')
            mask = valid_timestamps.notna()
            # Ensure the column is object type before assigning formatted strings
            display_recent_df['timestamp'] = display_recent_df['timestamp'].astype(object)
            display_recent_df.loc[mask, 'timestamp'] = valid_timestamps.loc[mask].dt.strftime("%Y-%m-%d %H:%M")
            
            # Add a selection column at the beginning
            display_recent_df.insert(0, 'Select', False)
            
            # Reset index for display but keep track of original indices
            display_recent_df = display_recent_df.reset_index(drop=True)
            
            # Create the editable dataframe with checkboxes
            edited_df = st.data_editor(
                display_recent_df,
                column_config={
                    "Select": st.column_config.CheckboxColumn(
                        "Select",
                        help="Select rows to delete",
                        default=False,
                    ),
                    "timestamp": st.column_config.TextColumn("Timestamp", disabled=True),
                    "action": st.column_config.TextColumn("Action", disabled=True),
                    "name": st.column_config.TextColumn("Name", disabled=True),
                    "company": st.column_config.TextColumn("Company", disabled=True),
                },
                disabled=["timestamp", "action", "name", "company"],
                hide_index=True,
                width='stretch',
                key="recent_activity_editor"
            )
            
            # Get selected rows
            selected_mask = edited_df['Select'] == True
            selected_count = selected_mask.sum()
            
            # Delete button section
            col1, col2 = st.columns([1, 4])
            with col1:
                delete_button = st.button(
                    f"🗑️ Delete Selected ({selected_count})",
                    disabled=selected_count == 0,
                    type="primary" if selected_count > 0 else "secondary",
                    key="delete_history_button"
                )
            
            with col2:
                if selected_count > 0:
                    st.caption(f"⚠️ {selected_count} record(s) selected for deletion")
            
            # Handle deletion
            if delete_button and selected_count > 0:
                # Get the original indices of selected rows
                selected_display_indices = edited_df[selected_mask].index.tolist()
                indices_to_delete = [original_indices[i] for i in selected_display_indices]
                
                # Confirm deletion
                success, message = delete_history_records(indices_to_delete)
                
                if success:
                    st.success(message)
                    # Rerun to refresh the data
                    st.rerun()
                else:
                    st.error(message)

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