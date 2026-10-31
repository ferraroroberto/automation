"""Regression tests for the DevTools orchestrator's Excel merge (temp dirs only)."""

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[4] / "linkedin/profiles_data_extractor/devtools"))

from linkedin_profiles_data_orchestrator_devtools import (  # noqa: E402
    LinkedInProfilesDataOrchestratorDevTools,
    WorkbookLockedError,
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


class LockedWorkbookTests(unittest.TestCase):
    """A workbook still locked after every retry is a failure, not 'profiles skipped'."""

    def setUp(self):
        self.orchestrator = LinkedInProfilesDataOrchestratorDevTools.__new__(
            LinkedInProfilesDataOrchestratorDevTools
        )
        self.profiles = [{"name": "Alex Example", "url_profile": "https://example.test/a"}]

    def _existing_workbook(self, tmp: str) -> str:
        destination = os.path.join(tmp, "out.xlsx")
        self.orchestrator.merge_to_excel([{"name": "Sam Example"}], destination)
        return destination

    def test_locked_before_reading_raises_and_names_the_unsaved_count(self):
        with tempfile.TemporaryDirectory() as tmp:
            destination = self._existing_workbook(tmp)
            with mock.patch.object(self.orchestrator, "wait_for_file_access", return_value=False):
                with self.assertRaises(WorkbookLockedError) as raised:
                    self.orchestrator.merge_to_excel(self.profiles, destination)
        self.assertIn("1 extracted profile(s) were NOT saved", str(raised.exception))

    def test_locked_just_before_writing_raises_and_leaves_the_file_untouched(self):
        with tempfile.TemporaryDirectory() as tmp:
            destination = self._existing_workbook(tmp)
            before = Path(destination).read_bytes()
            with mock.patch.object(self.orchestrator, "wait_for_file_access", side_effect=[True, False]):
                with self.assertRaises(WorkbookLockedError):
                    self.orchestrator.merge_to_excel(self.profiles, destination)
            self.assertEqual(Path(destination).read_bytes(), before)


if __name__ == "__main__":
    unittest.main()
