"""Dependencies live in the one root `requirements.txt`; secrets are read from `.env`, never a params file."""

from __future__ import annotations

import os
import re
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]

# The standalone PDF converter is the one tool `docs/layout.md` allows its own list.
ALLOWED_REQUIREMENTS = {"requirements.txt", "image/pdf_to_jpg/requirements.txt"}
SKIPPED_DIRS = {".git", ".venv", "node_modules", "__pycache__"}

# `params['api_token']` / `params.get("api_token")`: the legacy params-file token read.
PARAMS_TOKEN_READ = re.compile(r"""params(?:\[|\.get\()\s*['"]api_token['"]""")


def _walk(suffix: str) -> list[Path]:
    found = []
    for folder, dirs, files in os.walk(REPO_ROOT):
        dirs[:] = [d for d in dirs if d not in SKIPPED_DIRS]
        found.extend(Path(folder) / f for f in files if f.endswith(suffix))
    return found


class DependencyListTests(unittest.TestCase):
    def test_only_the_root_list_and_the_documented_exception_exist(self) -> None:
        listed = {
            p.relative_to(REPO_ROOT).as_posix()
            for p in _walk(".txt")
            if p.name.startswith("requirements")
        }
        self.assertEqual(listed - ALLOWED_REQUIREMENTS, set(), "add the pins to the root requirements.txt instead")


class SecretSourceTests(unittest.TestCase):
    def test_no_script_reads_the_api_token_from_a_params_file(self) -> None:
        offenders = [
            p.relative_to(REPO_ROOT).as_posix()
            for p in _walk(".py")
            if "tests" not in p.relative_to(REPO_ROOT).parts
            and PARAMS_TOKEN_READ.search(p.read_text(encoding="utf-8", errors="replace"))
        ]
        self.assertEqual(offenders, [], "read NOTION_API_TOKEN via utils.require_notion_api_token")


if __name__ == "__main__":
    unittest.main()
