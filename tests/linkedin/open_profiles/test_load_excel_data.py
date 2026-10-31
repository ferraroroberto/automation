"""A failed Excel read must surface as an exception the worker thread can report."""

import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "linkedin/open_profiles"))

import linkedin_open  # noqa: E402


class LoadExcelDataFailureTests(unittest.TestCase):
    def test_missing_file_raises_instead_of_exiting(self):
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(FileNotFoundError):
                linkedin_open.load_excel_data(str(Path(tmp, "absent.xlsx")))

    def test_locked_file_raises_a_permission_error_naming_the_file(self):
        with mock.patch.object(linkedin_open.pd, "read_excel", side_effect=PermissionError("locked")):
            with self.assertRaises(PermissionError) as raised:
                linkedin_open.load_excel_data("contacts.xlsx")
        self.assertIn("contacts.xlsx", str(raised.exception))

    def test_failure_is_an_exception_the_workers_except_branch_catches(self):
        # SystemExit is a BaseException; `except Exception` in process_data would not catch it.
        with mock.patch.object(linkedin_open.pd, "read_excel", side_effect=ValueError("bad")):
            try:
                linkedin_open.load_excel_data("x.xlsx")
            except Exception:
                return
        self.fail("no Exception raised")


if __name__ == "__main__":
    unittest.main()
