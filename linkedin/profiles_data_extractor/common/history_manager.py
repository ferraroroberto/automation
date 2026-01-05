import json
import os
from pathlib import Path
from datetime import datetime
import pandas as pd
import logging
import openpyxl

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

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

def ensure_history_file_exists():
    """Ensure the history file exists, creating it with headers if necessary."""
    history_path = get_history_file_path()
    if not history_path:
        logger.error("No history_file specified in config.")
        return False
        
    if not os.path.exists(history_path):
        try:
            # Create a new DataFrame with expected columns
            # We'll adapt columns as needed, but starting with standard ones plus timestamp
            df = pd.DataFrame(columns=[
                'timestamp', 'action', 'name', 'day', 'date connected', 
                'revocation_date', 'company', 'job_title'
            ])
            
            # Save empty dataframe
            with pd.ExcelWriter(history_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='History', index=False)
            logger.info(f"Created new history file at {history_path}")
            return True
        except Exception as e:
            logger.error(f"Failed to create history file: {e}")
            return False
            
    return True

def log_history(record_data, action="update"):
    """
    Log a record change to the history file.
    
    Args:
        record_data (dict): Dictionary containing the record data
        action (str): Type of action (update, create, delete, revoke, etc.)
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
        
        # Prepare new row
        new_row = record_data.copy()
        
        # Ensure timestamp is standard datetime
        new_row['timestamp'] = datetime.now()
        new_row['action'] = action
        
        # Create a DataFrame for the new row
        new_row_df = pd.DataFrame([new_row])
        
        # Concatenate with existing history
        updated_history_df = pd.concat([history_df, new_row_df], ignore_index=True)
        
        # Save back to Excel
        with pd.ExcelWriter(history_path, engine='openpyxl') as writer:
            updated_history_df.to_excel(writer, sheet_name='History', index=False)
            
        logger.info(f"Logged history for {record_data.get('name', 'Unknown')}")
        return True
        
    except Exception as e:
        logger.error(f"Failed to log history: {e}")
        return False
