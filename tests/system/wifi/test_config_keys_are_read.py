#!/usr/bin/env python3
"""Guard: every key in a Wi-Fi tool's JSON config and in its in-code fallback
default is read by the tool, and the two carry the same keys.

A knob nothing reads is a lie to whoever tunes it, and the fallback dict and
the JSON drift apart silently because each is only exercised when the other is
absent. This pins both against the script's own source.

Run from the repo root:

    & .\\.venv\\Scripts\\python.exe -m unittest discover -s tests/system/wifi -p "test_*.py"
"""

from __future__ import annotations

import json
import re
import sys
import unittest
from pathlib import Path
from typing import Any, Dict, Set
from unittest import mock

WIFI_DIR = Path(__file__).resolve().parents[3] / "system" / "wifi"
sys.path.insert(0, str(WIFI_DIR))

import network_scanner  # noqa: E402
import wifi_passwords  # noqa: E402

_TOOLS = (
    ("network_scanner", network_scanner),
    ("wifi_passwords", wifi_passwords),
)


def _flat_keys(node: Dict[str, Any], prefix: str = "") -> Set[str]:
    """Dotted paths of every leaf key (a nested block contributes its leaves, not itself)."""
    paths: Set[str] = set()
    for key, value in node.items():
        path = f"{prefix}{key}"
        if isinstance(value, dict):
            paths |= _flat_keys(value, f"{path}.")
        else:
            paths.add(path)
    return paths


class ConfigKeysAreReadTests(unittest.TestCase):
    def test_every_json_key_is_read_by_the_script(self) -> None:
        for name, _module in _TOOLS:
            source = (WIFI_DIR / f"{name}.py").read_text(encoding="utf-8")
            config = json.loads((WIFI_DIR / f"{name}.json").read_text(encoding="utf-8"))
            for path in sorted(_flat_keys(config)):
                leaf = path.rsplit(".", 1)[-1]
                read = re.search(rf"""(?:get\(|\[)\s*["']{re.escape(leaf)}["']""", source)
                self.assertIsNotNone(read, f"{name}.json key {path!r} is never read by {name}.py")

    def test_fallback_default_has_the_same_keys_as_the_json(self) -> None:
        for name, module in _TOOLS:
            config = json.loads((WIFI_DIR / f"{name}.json").read_text(encoding="utf-8"))
            with mock.patch.object(module.json, "load", side_effect=ValueError("forced fallback")):
                fallback = module.load_config()
            self.assertEqual(_flat_keys(fallback), _flat_keys(config), f"{name}: fallback and JSON keys differ")


if __name__ == "__main__":
    unittest.main()
