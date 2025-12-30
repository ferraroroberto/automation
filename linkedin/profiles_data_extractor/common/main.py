import streamlit as st
from pathlib import Path
import sys
import os

# Add the current directory to Python path to import local modules
current_dir = Path(__file__).parent
sys.path.insert(0, str(current_dir))

from dashboard import main as dashboard_main
from dataentry import main as dataentry_main

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

    # Create tabs for navigation
    tab1, tab2 = st.tabs(["📊 Dashboard", "✏️ Data Editor"])

    with tab1:
        st.header("Dashboard")
        dashboard_main()

    with tab2:
        st.header("Data Entry")
        dataentry_main()

if __name__ == "__main__":
    main()