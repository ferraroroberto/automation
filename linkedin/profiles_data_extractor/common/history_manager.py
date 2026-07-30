import json
import os
from pathlib import Path
from datetime import datetime, timedelta
import pandas as pd
import openpyxl

from _logging_setup import setup_logging

logger = setup_logging(__name__)

def load_config():
    """Load configuration from the JSON file in the same directory."""
    config_path = Path(__file__).parent / "linkedin_profiles_data.json"
    if not config_path.exists():
        logger.error(f"Config file not found at {config_path}")
        return None
    
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        logger.error(f"Error loading config: {e}")
        return None

def get_history_file_path():
    """Get the path to the history file from config."""
    config = load_config()
    if config:
        return config.get("history_file")
    return None

def migrate_history_columns(history_df):
    """Migrate existing history file to new column structure."""
    expected_columns = [
        'timestamp', 'action', 'date_contacted', 'search_type', 'url_search', 'name',
        'url_profile', 'job_title', 'follows_from', 'company', 'location',
        'reach_out_type', 'url_chat', 'date_connected', 'ind_answered', 'date_revocation',
        'date_discarded'
    ]

    # Add missing columns with None values
    for col in expected_columns:
        if col not in history_df.columns:
            history_df[col] = None

    # Reorder columns to match expected order
    history_df = history_df[expected_columns]

    return history_df

def ensure_history_file_exists():
    """Ensure the history file exists, creating it with headers if necessary."""
    history_path = get_history_file_path()
    if not history_path:
        logger.error("No history_file specified in config.")
        return False

    if not os.path.exists(history_path):
        try:
            # Create a new DataFrame with all base columns plus history columns
            # History columns come first: timestamp, action, then all base columns
            df = pd.DataFrame(columns=[
                'timestamp', 'action', 'date_contacted', 'search_type', 'url_search', 'name',
                'url_profile', 'job_title', 'follows_from', 'company', 'location',
                'reach_out_type', 'url_chat', 'date_connected', 'ind_answered', 'date_revocation',
                'date_discarded'
            ])

            # Save empty dataframe
            with pd.ExcelWriter(history_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='History', index=False)
            logger.info(f"Created new history file at {history_path}")
            return True
        except Exception as e:
            logger.error(f"Failed to create history file: {e}")
            return False
    else:
        # Check if existing file needs migration
        try:
            history_df = pd.read_excel(history_path, engine='openpyxl')
            original_columns = list(history_df.columns)

            expected_columns = [
                'timestamp', 'action', 'date_contacted', 'search_type', 'url_search', 'name',
                'url_profile', 'job_title', 'follows_from', 'company', 'location',
                'reach_out_type', 'url_chat', 'date_connected', 'ind_answered', 'date_revocation',
                'date_discarded'
            ]

            # Check if migration is needed
            if set(history_df.columns) != set(expected_columns):
                logger.info("Migrating history file to new column structure")
                history_df = migrate_history_columns(history_df)

                # Save migrated dataframe
                with pd.ExcelWriter(history_path, engine='openpyxl') as writer:
                    history_df.to_excel(writer, sheet_name='History', index=False)
                logger.info(f"Migrated history file columns from {original_columns} to {expected_columns}")

        except Exception as e:
            logger.error(f"Failed to migrate existing history file: {e}")
            return False

    return True

def log_history(record_data, action="update", original_data=None):
    """
    Log a record change to the history file.

    Args:
        record_data (dict): Dictionary containing the record data after changes
        action (str): Type of action (update, create, delete, revoke, etc.)
        original_data (dict, optional): Dictionary containing the record data before changes.
                                       Used for "initial" records when logging first-time updates.
    """
    history_path = get_history_file_path()
    if not history_path:
        logger.error("No history_file specified in config.")
        return False

    if not ensure_history_file_exists():
        return False

    try:
        # Load existing history
        history_df = pd.read_excel(history_path, engine='openpyxl')

        # Define all expected columns
        expected_columns = [
            'timestamp', 'action', 'date_contacted', 'search_type', 'url_search', 'name',
            'url_profile', 'job_title', 'follows_from', 'company', 'location',
            'reach_out_type', 'url_chat', 'date_connected', 'ind_answered', 'date_revocation',
            'date_discarded'
        ]

        # Ensure history_df has all expected columns
        for col in expected_columns:
            if col not in history_df.columns:
                history_df[col] = None

        # Reorder columns to match expected order
        history_df = history_df[expected_columns]

        # Check if this is the first time logging this record
        record_name = record_data.get('name')
        if record_name:
            existing_records = history_df[history_df['name'] == record_name]
            is_first_record = len(existing_records) == 0
        else:
            is_first_record = False

        # Get current timestamp for the change
        change_timestamp = datetime.now()

        # Prepare list of rows to add
        rows_to_add = []

        # If this is the first record for this name, add an "initial" entry first
        if is_first_record:
            # Use original_data if provided, otherwise use record_data (for backward compatibility)
            initial_data = original_data if original_data is not None else record_data
            initial_row = initial_data.copy()
            # Use a timestamp 1 second before the actual change for proper ordering
            initial_row['timestamp'] = change_timestamp - timedelta(seconds=1)
            initial_row['action'] = 'initial'

            # Ensure all expected columns are present in initial_row (fill missing with None)
            for col in expected_columns:
                if col not in initial_row:
                    initial_row[col] = None

            rows_to_add.append(initial_row)

        # Prepare the actual change row
        change_row = record_data.copy()
        # Use the change timestamp
        change_row['timestamp'] = change_timestamp
        change_row['action'] = action

        # Ensure all expected columns are present in change_row (fill missing with None)
        for col in expected_columns:
            if col not in change_row:
                change_row[col] = None

        rows_to_add.append(change_row)

        # Create DataFrames for all rows to add
        if rows_to_add:
            # Create new_rows_df ensuring it has the same structure as history_df
            new_rows_df = pd.DataFrame(rows_to_add, columns=expected_columns)

            # To avoid FutureWarning about empty/all-NA columns and ensure consistent dtypes,
            # we always convert all columns to object dtype before concatenation.
            # This ensures consistent behavior whether history_df is empty or not.
            for col in expected_columns:
                if col in history_df.columns:
                    history_df[col] = history_df[col].astype('object')
                if col in new_rows_df.columns:
                    new_rows_df[col] = new_rows_df[col].astype('object')
            
            updated_history_df = pd.concat([history_df, new_rows_df], ignore_index=True)

            # Save back to Excel
            with pd.ExcelWriter(history_path, engine='openpyxl') as writer:
                updated_history_df.to_excel(writer, sheet_name='History', index=False)

            action_desc = "initial + " + action if is_first_record else action
            logger.info(f"Logged history for {record_data.get('name', 'Unknown')} with action: {action_desc}")
        else:
            logger.warning("No rows to add to history")
            return False

        return True

    except Exception as e:
        logger.error(f"Failed to log history: {e}")
        return False
