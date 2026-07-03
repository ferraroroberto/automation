# requirements: public
import ast
import logging
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException
import win32com.client
import win32gui
from urllib.parse import unquote, urlsplit
import os
import pandas as pd
import pickle
from pathlib import Path
from datetime import datetime
from zipfile import BadZipFile
import json
from dotenv import load_dotenv
from typing import Optional, Dict, List, Any

log = logging.getLogger(__name__)


# Custom function to replace special characters
def replace_special_chars(path: str) -> str:
    special_char_mapping = {
        '%E1': 'á',
        '%E9': 'é',
        '%ED': 'í',
        '%F3': 'ó',
        '%FA': 'ú',
        '%C1': 'Á',
        '%C9': 'É',
        '%CD': 'Í',
        '%D3': 'Ó',
        '%DA': 'Ú',
        '%F1': 'ñ',
        '%D1': 'Ñ',
        '%20': ' '
    }

    for key, value in special_char_mapping.items():
        path = path.replace(key, value)

    return path

# Function to get the path of the foreground windows explorer
def get_first_explorer_hwnd() -> Optional[int]:
    # Get a list of all open windows
    windows = []
    win32gui.EnumWindows(lambda hwnd, windows: windows.append(hwnd), windows)

    # Find the first window that has a path to a directory in its title, taking into account network drives too
    for hwnd in windows:
        window_text = win32gui.GetWindowText(hwnd)
        if (":\\" in window_text or window_text.startswith("\\\\S555")) and not window_text.endswith(
                ".exe") and not window_text.endswith(".py"):
            return hwnd

    return None

def get_explorer_path_from_hwnd(target_hwnd: int) -> Optional[str]:
    # Get all instances of Shell Windows
    shell_windows = win32com.client.Dispatch("Shell.Application").Windows()

    # Filter for Windows Explorer instances
    explorer_windows = [w for w in shell_windows if w.LocationURL.startswith("file:///")]

    if not explorer_windows:
        log.warning("No Windows Explorer instances found.")
        return None

    # Iterate through explorer_windows to find the matching HWND and return the folder path
    for window in explorer_windows:
        hwnd = window.HWND
        if hwnd == target_hwnd:
            log.debug("LocationURL: %s", window.LocationURL)
            url_parts = urlsplit(window.LocationURL)
            folder_path = url_parts.path
            folder_path = folder_path[1:] if folder_path.startswith('/') else folder_path
            folder_path = folder_path.replace('/', '\\')
            folder_path = replace_special_chars(folder_path)
            return folder_path

    log.warning("No matching Windows Explorer instance found.")
    return None

def get_first_explorer_folder_path() -> Optional[str]:
    # Get the HWND of the first Windows Explorer instance with a path in its title
    first_explorer_hwnd = get_first_explorer_hwnd()

    # If no matching HWND is found, log an error message and return None
    if first_explorer_hwnd is None:
        log.warning("No Windows Explorer instance found with a path in its title.")
        return None

    # Get the folder path of the Windows Explorer instance with the matching HWND
    folder_path = get_explorer_path_from_hwnd(first_explorer_hwnd)

    # If a folder path is found, log the HWND and folder path, then return the folder path as a string
    if folder_path:
        log.info("Window handle: %s, Folder path: %s", first_explorer_hwnd, folder_path)
        return folder_path

    return None

# Default location of the legacy notion params txt file (single source of truth).
# Set the NOTION_PARAMS_FILE environment variable to the path on each machine.
DEFAULT_PARAMS_FILE = os.getenv("NOTION_PARAMS_FILE", "")

# Function to read the parameters from the txt file (legacy support)
def read_params_from_txt_file(file_path: str) -> Dict[str, str]:
    params = {}
    with open(file_path, 'r') as f:
        for line in f:
            if line.strip():
                key, value = line.strip().split(" = ", 1)
                params[key.strip()] = value.strip()
    return params

# Function to load JSON configuration file
def load_json_config(config_path: str) -> Dict[str, Any]:
    """
    Load configuration from a JSON file.

    Parameters:
    - config_path: Path to the JSON configuration file

    Returns:
    - Dictionary containing configuration data
    """
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        log.info("✅ Configuration loaded from %s", config_path)
        return config
    except FileNotFoundError:
        log.error("❌ Configuration file not found: %s", config_path)
        raise
    except json.JSONDecodeError as e:
        log.error("❌ Error parsing JSON configuration: %s", e)
        raise

# Function to load environment variables from .env file
def load_env_variables(env_path: Optional[str] = None) -> Dict[str, Optional[str]]:
    """
    Load environment variables from .env file.

    Parameters:
    - env_path: Optional path to .env file. If None, searches in current directory

    Returns:
    - Dictionary containing environment variables
    """
    if env_path:
        load_dotenv(env_path)
    else:
        load_dotenv()

    env_vars = {
        'notion_api_token': os.getenv('NOTION_API_TOKEN'),
        'notion_token_v2': os.getenv('NOTION_TOKEN_V2'),
        'notion_workspace_url': os.getenv('NOTION_WORKSPACE_URL')
    }

    log.info("✅ Environment variables loaded")
    return env_vars

# Function to open an excel file or a pickle, if found. If not found, creates the pickle
def read_excel_or_pickle(
    excel_file_path: str,
    pickle_file_path: str,
    sheet_name: Optional[str] = None,
    usecols: Optional[Any] = None,
    engine: Optional[str] = None,
) -> pd.DataFrame:
    excel_file = Path(excel_file_path)
    pickle_file = Path(pickle_file_path)

    start_time = datetime.now()
    log.info("Starting data loading process at %s", start_time)
    if pickle_file.exists() and pickle_file.stat().st_mtime > excel_file.stat().st_mtime:
        log.info("Loading data from pickle file %s", pickle_file)
        with open(pickle_file, 'rb') as f:
            df = pickle.load(f)
    else:
        log.info("Loading data from Excel file and creating a pickle file %s", excel_file)
        df = pd.read_excel(excel_file, sheet_name=sheet_name, usecols=usecols, engine=engine)
        with open(pickle_file, 'wb') as f:
            pickle.dump(df, f)

    end_time = datetime.now()
    log.info("Data loading process finished at %s", end_time)

    duration = end_time - start_time
    log.info("Total duration of the data loading: %s", duration)

    # If no sheet_name is specified and df is a dictionary, return the first DataFrame
    if sheet_name is None and isinstance(df, dict):
        first_sheet = list(df.values())[0]
        return first_sheet

    return df

# Get the column widths from the existing Excel file, initializing column_widths as an empty list first
def get_column_widths(excel_path: str) -> List[Any]:
    column_widths = []
    if not os.path.exists(excel_path):
        log.warning("No existing workbook at %s; skipping column width reuse.", excel_path)
        return column_widths

    try:
        wb_existing = openpyxl.load_workbook(excel_path)
        ws_existing = wb_existing.active

        # Find the last column with content
        last_col = ws_existing.max_column

        # Get column widths for columns with content
        column_widths = [
            ws_existing.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width
            for i in range(last_col)
        ]
    except (InvalidFileException, KeyError, BadZipFile) as exc:
        log.warning("Unable to read column widths from %s: %s", excel_path, exc)

    return column_widths

# Apply column widths to an excel file (requires the column_widths)
def apply_column_widths(excel_path: str, column_widths: List[Any]) -> None:
    if column_widths:
        wb_final = openpyxl.load_workbook(excel_path)
        ws_final = wb_final.active
        for i, width in enumerate(column_widths):
            ws_final.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width = width
        wb_final.save(excel_path)

def load_excel_with_json(excel_path: str, column_name: str) -> None:
    # Read the Excel file using the openpyxl engine
    df = pd.read_excel(excel_path, engine="openpyxl")

    # Define a function to load JSON nested dictionaries
    def load_json(json_str):
        # Convert the JSON string to a Python dictionary using ast.literal_eval()
        log.debug("Original JSON string: %s", json_str)
        return ast.literal_eval(json_str)

    # Apply the load_json() function to the specified column
    df[column_name] = df[column_name].apply(load_json)

    # Log the resulting DataFrame head
    log.debug("%s", df.head())

def load_excel_with_json_and_export(excel_path: str, column_name: str) -> None:
    # Read the Excel file using the openpyxl engine
    df = pd.read_excel(excel_path, engine="openpyxl")

    # Define a function to convert a JSON nested dictionary into a DataFrame
    def nested_dict_to_dataframe(nested_dict, prefix=''):
        """
        Converts a nested dictionary or list into a DataFrame with each key-value pair or list item in a separate column.
        The prefix argument is used to prefix the column names for nested dictionaries.
        """
        data = {}
        for key, value in nested_dict.items() if isinstance(nested_dict, dict) else enumerate(nested_dict):
            if isinstance(key, int):
                key = str(key)
            if isinstance(value, dict):
                # Recursively convert nested dictionaries into DataFrames
                data.update(nested_dict_to_dataframe(value, prefix=key + '_'))
            elif isinstance(value, list):
                # Treat each list as a separate column with column names based on the index of the list item
                for i, item in enumerate(value):
                    data[prefix + key + '_' + str(i)] = item
            else:
                # Add the key-value pair to the data dictionary
                data[prefix + key] = value
        return pd.DataFrame(data, index=[0])

    # Define a function to load JSON nested dictionaries
    def load_json(json_str):
        # Convert the JSON string to a Python dictionary using ast.literal_eval()
        return ast.literal_eval(json_str)

    # Apply the load_json() function to the specified column
    df[column_name] = df[column_name].apply(load_json)

    # Create a new DataFrame with the structure of the JSON nested dictionaries
    dict_df = pd.concat([nested_dict_to_dataframe(row, prefix=column_name+'_') for row in df[column_name]], ignore_index=True)

    # Save the DataFrame as an Excel file with the same path as the input file and a suffix
    output_path = excel_path[:-5] + '_' + column_name + '_dictionary.xlsx'
    dict_df.to_excel(output_path, index=False)

    # Log the resulting DataFrame head
    log.debug("%s", dict_df.head())
