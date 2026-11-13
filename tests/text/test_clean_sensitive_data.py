"""The cleaner redacts URLs and emails whole before the path pass can split them."""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tools" / "text"))

import clean_sensitive_data  # noqa: E402


class CleanSensitiveDataTests(unittest.TestCase):
    def setUp(self):
        self.cleaner = clean_sensitive_data.DataCleanerBase()

    def test_url_with_query_string_is_redacted_whole(self):
        cleaned = self.cleaner.clean_sensitive_data("see https://host/report?customer=alice now")
        self.assertEqual(cleaned, "see [REDACTED_URL] now")

    def test_windows_path_is_still_redacted(self):
        cleaned = self.cleaner.clean_sensitive_data(r"opened C:\Users\alice\notes.txt today")
        self.assertEqual(cleaned, "opened [REDACTED_PATH] today")

    def test_email_is_redacted(self):
        cleaned = self.cleaner.clean_sensitive_data("mail alice@example.com please")
        self.assertEqual(cleaned, "mail [REDACTED_EMAIL] please")


if __name__ == "__main__":
    unittest.main()
