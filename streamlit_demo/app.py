"""
Streamlit Demo App – Main Entry Point
======================================
Launch with:
    streamlit run app.py

This file configures the page, renders the sidebar menu, and dynamically
dispatches to the selected demo module.  Each demo lives in its own file
under the ``pages/`` package and exposes a ``render()`` function.
"""

import streamlit as st

from menu import PAGES, render_home

# ---------------------------------------------------------------------------
# Page configuration (must be the first Streamlit call)
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Streamlit Demo Playground",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Sidebar navigation
# ---------------------------------------------------------------------------
st.sidebar.title("Streamlit Demo")
st.sidebar.markdown("---")

# Build the list of page labels (Home + every registered demo)
page_labels = ["Home"] + [p["label"] for p in PAGES]

selection = st.sidebar.radio(
    "Navigate to",
    page_labels,
    index=0,
    key="nav_radio",
)

st.sidebar.markdown("---")
st.sidebar.caption("Built for internal training purposes.")

# ---------------------------------------------------------------------------
# Dispatch to the selected page
# ---------------------------------------------------------------------------
if selection == "Home":
    render_home()
else:
    # Find the matching page dict and call its render function
    for page in PAGES:
        if page["label"] == selection:
            page["module"].render()
            break
