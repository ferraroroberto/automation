"""The data-entry and reach-out pages share one re-contact rule, one type-option list and one config reader."""

import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import pandas as pd

os.environ.setdefault("LINKEDIN_EXTRACTOR_LOG_FILE", os.path.join(tempfile.gettempdir(), "linkedin_extractor_tests.log"))
sys.path.insert(0, str(Path(__file__).resolve().parents[4] / "linkedin/profiles_data_extractor/common"))

import _lib  # noqa: E402
import profiles_config  # noqa: E402

REVOKED = pd.Timestamp("2024-05-01")
OLD_DAY = pd.Timestamp("2024-01-01")
NEW_DAY = pd.Timestamp("2025-02-03")


class RecontactRuleTests(unittest.TestCase):
    def apply(self, new_day, original_day, revocation):
        with mock.patch.object(_lib.st, "info") as info:
            return _lib.apply_recontact_rule(new_day, original_day, revocation), info

    def test_a_new_contact_date_clears_the_revocation_and_says_so(self) -> None:
        for original in (OLD_DAY, pd.NaT, None):
            with self.subTest(original=original):
                revocation, info = self.apply(NEW_DAY, original, REVOKED)
                self.assertIsNone(revocation)
                info.assert_called_once()

    def test_a_new_contact_date_stays_silent_when_nothing_was_revoked(self) -> None:
        revocation, info = self.apply(NEW_DAY, OLD_DAY, pd.NaT)
        self.assertIsNone(revocation)
        info.assert_not_called()

    def test_an_unchanged_or_cleared_contact_date_keeps_the_revocation(self) -> None:
        for new_day in (OLD_DAY, pd.NaT, None):
            with self.subTest(new_day=new_day):
                revocation, info = self.apply(new_day, OLD_DAY, REVOKED)
                self.assertEqual(revocation, REVOKED)
                info.assert_not_called()


class ReachOutTypeOptionsTests(unittest.TestCase):
    def test_options_are_an_empty_choice_then_the_sorted_distinct_types(self) -> None:
        df = pd.DataFrame({"reach_out_type": ["Warm", None, "Cold", "Warm"]})
        self.assertEqual(_lib.reach_out_type_options(df), ["", "Cold", "Warm"])

    def test_the_records_own_type_is_always_offered(self) -> None:
        df = pd.DataFrame({"reach_out_type": ["Warm"]})
        self.assertEqual(_lib.reach_out_type_options(df, "Event"), ["", "Event", "Warm"])

    def test_a_frame_without_the_column_offers_only_the_empty_choice(self) -> None:
        self.assertEqual(_lib.reach_out_type_options(pd.DataFrame({"name": ["Ana"]})), [""])


class ProfilesConfigTests(unittest.TestCase):
    def test_existing_names_are_lowercased_stripped_and_non_empty(self) -> None:
        df = pd.DataFrame({"name": ["  Ana Sample ", "BOB", None, "   ", 7]})
        self.assertEqual(profiles_config.existing_contact_names(df), {"ana sample", "bob"})
        self.assertEqual(profiles_config.existing_contact_names(pd.DataFrame({"x": [1]})), set())

    def test_the_loader_parses_a_given_path_and_raises_when_it_is_missing(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "config.json"
            path.write_text(json.dumps({"destination_file": "x.xlsx"}), encoding="utf-8")
            self.assertEqual(profiles_config.load_profiles_config(path), {"destination_file": "x.xlsx"})
            with self.assertRaises(FileNotFoundError):
                profiles_config.load_profiles_config(Path(tmp) / "absent.json")


if __name__ == "__main__":
    unittest.main()
