"""Regression tests for the DevTools orchestrator's Excel merge (temp dirs only)."""

import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from linkedin_profiles_data_orchestrator_devtools import (  # noqa: E402
    LinkedInProfilesDataOrchestratorDevTools,
)


class MergeToExcelReturnShapeTests(unittest.TestCase):
    def setUp(self):
        self.orchestrator = LinkedInProfilesDataOrchestratorDevTools.__new__(
            LinkedInProfilesDataOrchestratorDevTools
        )

    def test_no_named_profile_returns_the_three_values_the_caller_unpacks(self):
        # The caller does `inserted, skipped, total = self.merge_to_excel(...)`;
        # the no-valid-profile path returned a 2-tuple and crashed it.
        with tempfile.TemporaryDirectory() as tmp:
            destination = os.path.join(tmp, "out.xlsx")
            result = self.orchestrator.merge_to_excel(
                [{"name": "  ", "url_profile": "https://example.test/a"}, {"job_title": "x"}],
                destination,
            )
            inserted, skipped, total = result
            self.assertEqual((inserted, skipped, total), (0, 0, 0))
            self.assertFalse(os.path.exists(destination))

    def test_new_file_path_returns_three_values(self):
        with tempfile.TemporaryDirectory() as tmp:
            destination = os.path.join(tmp, "out.xlsx")
            inserted, skipped, total = self.orchestrator.merge_to_excel(
                [{"name": "Alex Example", "url_profile": "https://example.test/a"}], destination
            )
            self.assertEqual((inserted, skipped, total), (1, 0, 1))


if __name__ == "__main__":
    unittest.main()
