"""
Data Loader
============
Centralised helper that loads mock data for every demo page.
If the CSV / JSON files do not exist yet it triggers generation
automatically so the app always starts cleanly.
"""

import sys
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# Paths (resolved relative to *this* file so imports work from anywhere)
# ---------------------------------------------------------------------------
_DATA_DIR = Path(__file__).resolve().parent
MOCK_DIR = _DATA_DIR / "mock_data"

EMPLOYEES_CSV = MOCK_DIR / "employees.csv"
EMPLOYEES_JSON = MOCK_DIR / "employees.json"
SALES_CSV = MOCK_DIR / "sales.csv"


def _ensure_mock_data() -> None:
    """Generate mock data on-the-fly if files are missing."""
    if EMPLOYEES_CSV.exists() and SALES_CSV.exists():
        return
    # Add the project root so the scripts package is importable
    project_root = _DATA_DIR.parent
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))
    from scripts.generate_mock_data import generate_all
    generate_all()


def load_employees() -> pd.DataFrame:
    """Return the employees dataset as a DataFrame."""
    _ensure_mock_data()
    return pd.read_csv(EMPLOYEES_CSV)


def load_sales() -> pd.DataFrame:
    """Return the sales dataset as a DataFrame."""
    _ensure_mock_data()
    return pd.read_csv(SALES_CSV)
