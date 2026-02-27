"""
Central navigation and page registry for the Streamlit demo app.

Each demo page is registered here as a callable. The main app uses this
registry to build the sidebar navigation and dynamically render pages.
"""

import importlib
from dataclasses import dataclass
from typing import Callable

PAGE_MODULE_PREFIX = "pages"

@dataclass
class DemoPage:
    """Metadata for a single demo page."""
    key: str
    title: str
    icon: str
    description: str
    module_name: str  # relative to the pages package


# Ordered list of all demo pages shown in the sidebar.
PAGE_REGISTRY: list[DemoPage] = [
    DemoPage(
        key="home",
        title="Home",
        icon="🏠",
        description="Welcome page and navigation hub.",
        module_name="",
    ),
    DemoPage(
        key="data_input",
        title="Data Input",
        icon="📝",
        description="Text inputs, sliders, selectboxes, forms, and submit buttons.",
        module_name="data_input",
    ),
    DemoPage(
        key="visualization",
        title="Data Visualization",
        icon="📊",
        description="Tables, editable dataframes, charts, and KPI metrics.",
        module_name="visualization",
    ),
    DemoPage(
        key="crud_demo",
        title="CRUD Operations",
        icon="🗃️",
        description="Create, read, update, and delete records from an in-memory dataset.",
        module_name="crud_demo",
    ),
    DemoPage(
        key="file_upload",
        title="File Handling",
        icon="📂",
        description="Upload CSV/TXT/JSON files and download generated files.",
        module_name="file_upload",
    ),
    DemoPage(
        key="process_runner",
        title="Process Runner",
        icon="⚙️",
        description="Execute a simulated long-running process with live logs and progress.",
        module_name="process_runner",
    ),
    DemoPage(
        key="state_management",
        title="State Management",
        icon="🧠",
        description="Demonstrate st.session_state and cross-page state sharing.",
        module_name="state_management",
    ),
]


def load_page(page: DemoPage) -> Callable:
    """Dynamically import a demo page module and return its `render` function."""
    module = importlib.import_module(f"{PAGE_MODULE_PREFIX}.{page.module_name}")
    return module.render
