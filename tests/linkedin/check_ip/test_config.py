"""Regression tests for check_ip's --config handling (temp config files only)."""

import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "linkedin/check_ip"))

import config as check_ip_config  # noqa: E402


def _write_config(folder: Path, name: str, serpapi_key: str) -> Path:
    data = {
        "api_keys": {
            "imgur_client_id": "id",
            "imgur_access_token": "token",
            "serpapi_key": serpapi_key,
        },
        "folders": {"images_folder": "./images", "metadata_folder": "./metadata"},
        "file_names": {
            "metadata_file": "metadata.xlsx",
            "results_file": "results.xlsx",
            "api_history_file": "history.xlsx",
        },
        "processing_thresholds": {
            "linkedin_high_threshold": 50,
            "high_linkedin_days": 7,
            "standard_days": 180,
        },
        "search_settings": {"do_exact_search": True, "do_similar_search": False},
        "allowed_extensions": [".jpg"],
        "api_endpoints": {"search_endpoint": "https://example.test/search"},
    }
    path = folder / name
    path.write_text(json.dumps(data), encoding="utf-8")
    return path


class CustomConfigFlagTests(unittest.TestCase):
    def _parse(self, script_dir: Path, *argv: str):
        # Config() resolves its default file next to config.py; point that at a
        # temp folder so the real default file is never read or created.
        with mock.patch.object(check_ip_config, "__file__", str(script_dir / "config.py")), \
                mock.patch.object(sys, "argv", ["main.py", *argv]):
            return check_ip_config.parse_arguments()

    def test_custom_config_supplies_the_argparse_defaults(self):
        with tempfile.TemporaryDirectory() as script_dir, tempfile.TemporaryDirectory() as other:
            _write_config(Path(script_dir), "linkedin_search_image.json", "default-key")
            custom = _write_config(Path(other), "custom.json", "custom-key")
            args = self._parse(Path(script_dir), "--config", str(custom))
            self.assertEqual(args.serpapi_key, "custom-key")

    def test_default_config_is_still_used_without_the_flag(self):
        with tempfile.TemporaryDirectory() as script_dir:
            _write_config(Path(script_dir), "linkedin_search_image.json", "default-key")
            args = self._parse(Path(script_dir))
            self.assertEqual(args.serpapi_key, "default-key")

    def test_custom_config_works_when_the_default_file_is_missing(self):
        with tempfile.TemporaryDirectory() as script_dir, tempfile.TemporaryDirectory() as other:
            custom = _write_config(Path(other), "custom.json", "custom-key")
            args = self._parse(Path(script_dir), "--config", str(custom))
            self.assertEqual(args.serpapi_key, "custom-key")

    def test_command_line_flag_still_overrides_the_custom_config(self):
        with tempfile.TemporaryDirectory() as script_dir, tempfile.TemporaryDirectory() as other:
            custom = _write_config(Path(other), "custom.json", "custom-key")
            args = self._parse(Path(script_dir), "--config", str(custom), "--serpapi-key", "cli-key")
            self.assertEqual(args.serpapi_key, "cli-key")


class MissingConfigTests(unittest.TestCase):
    def test_missing_config_exits_without_writing_a_default_file(self):
        with tempfile.TemporaryDirectory() as folder:
            missing = Path(folder) / "linkedin_search_image.json"
            with self.assertRaises(SystemExit) as raised:
                check_ip_config.Config(str(missing))
            self.assertNotEqual(raised.exception.code, 0)
            self.assertFalse(missing.exists())


if __name__ == "__main__":
    unittest.main()
