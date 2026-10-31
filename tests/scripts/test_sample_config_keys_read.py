"""Every key a shipped sample config advertises must be read by the tool's code.

A sample key nothing reads is a setting a user edits expecting an effect and
gets none. This walks each tracked `*.json.sample` / `*.example.json`, collects
its object keys, and fails when a key never appears as a quoted string in the
tracked `.py` / `.ps1` / `.js` sources of the tool that owns the sample. There is
no baseline: a dead key is deleted from the loader, the sample and the doc
together, or the setting is wired up.
"""

from __future__ import annotations

import json
import subprocess
import sys
import unittest
from pathlib import Path
from typing import Iterator

REPO_ROOT = Path(__file__).resolve().parents[2]

# Repo-root shared helper.
sys.path.append(str(REPO_ROOT))
from no_window import NO_WINDOW  # noqa: E402
SOURCE_SUFFIXES = (".py", ".ps1", ".js")

# Sample files whose keys are user data or a device dump, not settings.
DATA_SAMPLES = {
    "smart_life/devices.sample.json",
    "system/textexpander/text_expander_config.example.json",
}

# Keys that are user-chosen names (counter labels used as {placeholders}), not settings.
USER_NAMED_KEYS = {"google/photos_weekly_album.json.sample": {"first", "second"}}

# Self-describing metadata no code is expected to read.
METADATA_KEYS = {"_comment", "author", "created_at", "description", "module_name", "version"}


def _tracked_files() -> list[str]:
    out = subprocess.run(
        ["git", "ls-files"], cwd=REPO_ROOT, capture_output=True, text=True, check=True, creationflags=NO_WINDOW
    ).stdout
    return out.splitlines()


def _is_sample(path: str) -> bool:
    name = path.lower()
    return name.endswith(".json.sample") or name.endswith(".example.json") or name.endswith(".sample.json")


def _keys(node: object) -> Iterator[str]:
    if isinstance(node, dict):
        for key, value in node.items():
            yield key
            yield from _keys(value)
    elif isinstance(node, list):
        for item in node:
            yield from _keys(item)


def _tool_scope(sample: str) -> Path:
    # A sample at <domain>/<tool>/... is read by code anywhere under that tool;
    # one directly in <domain>/ is read by its sibling scripts.
    parts = Path(sample).parts[:-1]
    return Path(*parts[:2])


def dead_keys(sample: str, tracked: list[str]) -> list[str]:
    scope = _tool_scope(sample)
    source = "".join(
        (REPO_ROOT / f).read_text(encoding="utf-8", errors="ignore")
        for f in tracked
        if f.endswith(SOURCE_SUFFIXES) and Path(f).is_relative_to(scope)
    )
    data = json.loads((REPO_ROOT / sample).read_text(encoding="utf-8"))
    return sorted(
        key
        for key in set(_keys(data))
        if key not in METADATA_KEYS
        and key not in USER_NAMED_KEYS.get(sample, ())
        and f'"{key}"' not in source and f"'{key}'" not in source
    )


class SampleConfigKeysAreReadTests(unittest.TestCase):
    def test_every_sample_key_is_read_by_its_tool(self):
        tracked = _tracked_files()
        failures = {}
        for sample in filter(_is_sample, tracked):
            if sample in DATA_SAMPLES:
                continue
            dead = dead_keys(sample, tracked)
            if dead:
                failures[sample] = dead
        self.assertEqual(failures, {}, "sample keys that no code reads")


# docs/layout.md names requirements.txt the one dependency list; this converter
# is the documented standalone exception.
STANDALONE_REQUIREMENTS = {"image/pdf_to_jpg/requirements.txt"}


class NoOrphanRequirementsTests(unittest.TestCase):
    def test_only_the_root_requirements_file_exists_besides_documented_exceptions(self):
        extra = sorted(
            f for f in _tracked_files()
            if Path(f).name == "requirements.txt" and f != "requirements.txt" and f not in STANDALONE_REQUIREMENTS
        )
        self.assertEqual(extra, [], "per-tool requirements.txt that nothing installs from")


if __name__ == "__main__":
    unittest.main()
