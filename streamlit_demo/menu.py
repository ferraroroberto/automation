"""
Menu / Page Registry
====================
Central registry of every demo page.  Each entry maps a human-readable
label and short description to the module that implements the demo.

Adding a new demo is as simple as:
1. Create ``pages/my_demo.py`` with a ``render()`` function.
2. Import it here and append an entry to ``PAGES``.
"""

import streamlit as st

# ---------------------------------------------------------------------------
# Import every demo module (each exposes a ``render()`` callable)
# ---------------------------------------------------------------------------
from pages import (
    crud_demo,
    data_input,
    file_upload,
    process_runner,
    state_management,
    visualization,
)

# ---------------------------------------------------------------------------
# Page registry – order here = order in the sidebar
# ---------------------------------------------------------------------------
PAGES: list[dict] = [
    {
        "label": "Data Input",
        "icon": "⌨",
        "description": "Text fields, sliders, selects, forms with submit buttons.",
        "module": data_input,
    },
    {
        "label": "Visualization",
        "icon": "📊",
        "description": "Tables, editable dataframes, line/bar/scatter charts, KPI metrics.",
        "module": visualization,
    },
    {
        "label": "CRUD Operations",
        "icon": "🗂",
        "description": "Create, read, update and delete records from an in-memory dataset.",
        "module": crud_demo,
    },
    {
        "label": "File Handling",
        "icon": "📁",
        "description": "Upload CSV / TXT / JSON files and download generated files.",
        "module": file_upload,
    },
    {
        "label": "Process Runner",
        "icon": "💻",
        "description": "Execute a long-running task with live logs, progress bars and status.",
        "module": process_runner,
    },
    {
        "label": "State Management",
        "icon": "🧠",
        "description": "Demonstrate st.session_state and cross-module shared state.",
        "module": state_management,
    },
]


# ---------------------------------------------------------------------------
# Home page renderer
# ---------------------------------------------------------------------------
def render_home() -> None:
    """Render the landing / home page with an overview of all demos."""
    st.title("Streamlit Demo Playground")
    st.markdown(
        """
        Welcome! This application is a **self-contained reference project** that
        showcases the most common Streamlit patterns you will need when building
        internal tools and dashboards.

        Use the **sidebar** on the left to navigate between demos.  Each demo is
        a standalone module that focuses on a single capability.
        """
    )

    st.markdown("---")
    st.subheader("Available Demos")

    # Render a card-like grid of demos
    cols = st.columns(3)
    for idx, page in enumerate(PAGES):
        with cols[idx % 3]:
            st.markdown(f"#### {page['icon']} {page['label']}")
            st.write(page["description"])

    st.markdown("---")
    st.info(
        "**Tip:** All data is generated locally — no external services are required. "
        "If the mock data files are missing they will be created automatically on first run."
    )
