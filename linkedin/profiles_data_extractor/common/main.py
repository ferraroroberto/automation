import streamlit as st
from pathlib import Path
import sys
import os
import json
import pandas as pd

# Add the current directory to Python path to import local modules
current_dir = Path(__file__).parent
sys.path.insert(0, str(current_dir))

from dashboard import main as dashboard_main
from dataentry import main as dataentry_main
from reachout import main as reachout_main

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

def main():
    """Main navigation hub for LinkedIn Reachout Dashboard."""

    # Set page config
    st.set_page_config(
        page_title="LinkedIn Reachout Hub",
        page_icon="🔗",
        layout="wide"
    )

    # Reduce spacing before title
    st.markdown("<style>.block-container { padding-top: 1.5rem; }</style>", unsafe_allow_html=True)

    st.title("🔗 LinkedIn Reachout Hub")

    # Load config and data once for shared use
    config = load_config()
    if not config:
        return

    data_path = config.get("destination_file")
    if not data_path:
        st.error("No 'destination_file' specified in config.")
        return

    df = load_excel_data(data_path)
    if df is None:
        return

    # Shared Sidebar Filters
    st.sidebar.header("Filters")

    # Note about Data Editor
    st.sidebar.info("ℹ️ Filters apply to Dashboard and Reachout tabs only. Data Editor is not affected by these filters.")

    # Company Filter
    df_filtered = df.copy()
    if 'company' in df_filtered.columns:
        companies = ['All'] + sorted(df_filtered['company'].dropna().unique().tolist())
        selected_company = st.sidebar.selectbox("Select Company", companies)
        if selected_company != 'All':
            df_filtered = df_filtered[df_filtered['company'] == selected_company]

    # Search Type Filter
    if 'search_type' in df_filtered.columns:
        search_types = ['All'] + sorted(df_filtered['search_type'].dropna().unique().tolist())
        selected_type = st.sidebar.selectbox("Select Search Type", search_types)
        if selected_type != 'All':
            df_filtered = df_filtered[df_filtered['search_type'] == selected_type]

    # Create tabs for navigation
    tab1, tab2, tab3 = st.tabs(["📊 Dashboard", "✏️ Data Editor", "🎯 Reachout"])

    with tab1:
        st.header("Dashboard")
        dashboard_main(df_filtered, df)  # Pass both filtered and full datasets

    with tab2:
        st.header("Data Entry")
        dataentry_main()

    with tab3:
        st.header("Reachout")
        reachout_main(df_filtered, df)  # Pass both filtered and full datasets

if __name__ == "__main__":
    main()