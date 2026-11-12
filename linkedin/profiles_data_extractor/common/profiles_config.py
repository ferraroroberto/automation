"""
Streamlit-free access to ``linkedin_profiles_data.json`` and the contacts workbook.

The Streamlit pages (through ``loaders``), the history manager and the devtools
scripts all read the same config and compare profile names the same way, so the
path, the loader and the existing-names rule each live here once.
"""

import json
from pathlib import Path
from typing import Optional, Set

import pandas as pd

CONFIG_PATH = Path(__file__).parent / "linkedin_profiles_data.json"


def load_profiles_config(config_path: Optional[Path] = None) -> dict:
    """Parse the JSON config (the shared one unless a path is given).

    Raises ``FileNotFoundError`` when the file is missing and ``ValueError``
    (``json.JSONDecodeError`` included) when it is not valid JSON; callers
    decide how to report either.
    """
    path = Path(config_path) if config_path else CONFIG_PATH
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def existing_contact_names(df: pd.DataFrame) -> Set[str]:
    """Lowercased, stripped, non-empty names of the ``name`` column, for case-insensitive de-duplication."""
    if "name" not in df.columns:
        return set()
    return {
        name.lower().strip()
        for name in df["name"].dropna()
        if isinstance(name, str) and name.strip()
    }
