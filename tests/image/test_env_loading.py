"""Entry points that read variables documented in `.env.sample` must load `.env`.

Each case imports a script by path with `dotenv.load_dotenv` replaced by a fake
that "loads" the variable, then checks the script picked it up. The real `.env`
is never read.
"""

import importlib.util
import os
import sys
import types
import unittest
from pathlib import Path
from unittest import mock

REPO_ROOT = Path(__file__).resolve().parents[2]

# (script, variable, how to read the value the script derived from it)
CASES = [
    ("image/carrousel_pdf_to_jpg.py", "CARROUSEL_SOURCE_FOLDER", lambda m: m.HARDCODED_SOURCE_FOLDER),
    ("image/carrousel_pdf_to_jpg.py", "POPPLER_PATH", lambda m: m.POPPLER_PATH),
    ("image/collage_image.py", "COLLAGE_SOURCE_FOLDER", lambda m: m.source_folder),
    (
        "image/illustrations_formatter.py",
        "ILLUSTRATIONS_DEST_INSTAGRAM",
        lambda m: m.default_illustrations_config()["destination_folder_instagram"],
    ),
    (
        "image/illustrations_formatter.py",
        "ILLUSTRATIONS_DEST_1920X1080",
        lambda m: m.default_illustrations_config()["destination_folder_1920x1080"],
    ),
    (
        "image/illustrations_formatter_gui.py",
        "ILLUSTRATIONS_DEST_INSTAGRAM",
        lambda m: m.IllustrationsFormatterGUI.DEFAULT_INSTAGRAM_FOLDER,
    ),
    (
        "image/illustrations_formatter_gui.py",
        "ILLUSTRATIONS_DEST_1920X1080",
        lambda m: m.IllustrationsFormatterGUI.DEFAULT_1920X1080_FOLDER,
    ),
    ("notion/todoist_migration.py", "TODOIST_SOURCE_FOLDER", lambda m: m.HARDCODED_SOURCE_FOLDER),
]


def _value_after_import(relative_path: str, variable: str, read):
    """Import `relative_path` with a fake load_dotenv that sets `variable`, then `read` it."""
    path = REPO_ROOT / relative_path

    def fake_load_dotenv(*args, **kwargs):
        os.environ[variable] = "from-dotenv"
        return True

    # pdf2image (and poppler) are not needed to read the module-level settings.
    stub_pdf2image = types.ModuleType("pdf2image")
    stub_pdf2image.convert_from_path = lambda *args, **kwargs: []

    with mock.patch.dict(os.environ), mock.patch.dict(sys.modules), \
            mock.patch("dotenv.load_dotenv", fake_load_dotenv), \
            mock.patch.object(sys, "path", [str(path.parent), *sys.path]):
        os.environ.pop(variable, None)
        sys.modules["pdf2image"] = stub_pdf2image
        spec = importlib.util.spec_from_file_location(f"envcheck_{path.stem}", path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        # Read inside the patched environment: some settings are resolved lazily.
        return read(module)


class EnvLoadingTests(unittest.TestCase):
    def test_scripts_pick_up_variables_from_dotenv(self):
        for relative_path, variable, read in CASES:
            with self.subTest(script=relative_path, variable=variable):
                self.assertEqual(_value_after_import(relative_path, variable, read), "from-dotenv")


if __name__ == "__main__":
    unittest.main()
