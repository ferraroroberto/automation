"""
Streamlit Demo App – Main Entry Point
======================================
Launch with:
    streamlit run main_menu.py

This file configures the page and renders the Home/landing content.
The individual demos are exposed as Streamlit multipage entries via the
``pages/`` directory.
"""

import streamlit as st

from menu import render_home

# ---------------------------------------------------------------------------
# Page configuration (must be the first Streamlit call)
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Streamlit Demo Playground",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded",
)

render_home()
