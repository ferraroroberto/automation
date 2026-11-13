"""The single-profile and search-results scripts must share one card parser."""

import os
import sys
import tempfile
import unittest
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[4]
os.environ.setdefault("LINKEDIN_EXTRACTOR_LOG_FILE", os.path.join(tempfile.gettempdir(), "linkedin_extractor_tests.log"))
sys.path.insert(0, str(_ROOT / "tools/linkedin/profiles_data_extractor/devtools"))

import linkedin_profiles_data_extractor_devtools as extractor  # noqa: E402


class SharedCardParserTests(unittest.TestCase):
    def test_both_scripts_embed_the_one_card_fragment(self):
        self.assertIn(extractor._CARD_JS, extractor._PROFILE_JS)
        self.assertIn(extractor._CARD_JS, extractor._SEARCH_RESULTS_JS)

    def test_company_selector_list_is_defined_once(self):
        for script in (extractor._PROFILE_JS, extractor._SEARCH_RESULTS_JS):
            self.assertEqual(script.count("const companySelectors"), 1)
            self.assertIn("search-reusables__filter-pill-button", script)

    def test_url_cleanup_is_defined_once(self):
        for script in (extractor._PROFILE_JS, extractor._SEARCH_RESULTS_JS):
            self.assertEqual(script.count("href.startsWith('/in/')"), 1)


if __name__ == "__main__":
    unittest.main()
