"""
Shared I/O helpers for the LinkedIn profiles data extractor.

All Streamlit pages inside this package import ``load_config`` and
``load_excel_data`` from here so the loading logic lives in exactly one place.
"""

import os
from typing import Optional

import pandas as pd
import streamlit as st
from profiles_config import CONFIG_PATH, load_profiles_config


def load_config() -> Optional[dict]:
    """Load configuration from the JSON file in the same directory.

    Returns the parsed JSON dict, or ``None`` on any error (displays
    ``st.error`` in the Streamlit UI before returning).
    """
    try:
        return load_profiles_config()
    except FileNotFoundError:
        st.error(f"Config file not found at {CONFIG_PATH}")
    except (OSError, ValueError) as e:
        st.error(f"Error loading config: {e}")
    return None


def load_excel_data(file_path: str) -> Optional[pd.DataFrame]:
    """Load data from an Excel file and normalise all date columns to datetime.

    Creates any missing expected date columns as ``pd.NaT`` so callers can
    always assume those columns exist.  Returns ``None`` on any error.
    """
    if not os.path.exists(file_path):
        st.error(f"Data file not found at: {file_path}")
        return None

    try:
        df = pd.read_excel(file_path)

        date_columns = ["date_contacted", "date_connected", "date_revocation", "date_discarded"]
        for col in date_columns:
            if col not in df.columns:
                df[col] = pd.NaT
            df[col] = pd.to_datetime(df[col], errors="coerce")

        return df
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return None
