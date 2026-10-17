"""Characterization tests for ``update_existing_record`` (in-memory frames only).

Both Streamlit pages edit a record through this function. It used to exist as
two diverging copies (``dataentry`` and ``reachout``); the cases below pin the
results every real caller sees, and run against whichever object each page
exposes.
"""

import datetime
import os
import sys
import unittest

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import _lib  # noqa: E402
import dataentry  # noqa: E402
import reachout  # noqa: E402

PAGE_IMPLEMENTATIONS = {
    "dataentry": dataentry.update_existing_record,
    "reachout": reachout.update_existing_record,
}


def make_frame():
    """Shape of ``loaders.load_excel_data``: date columns are always datetime64."""
    return pd.DataFrame({
        "name": ["Ana Sample", "Bob Example"],
        "company": ["Acme", "Globex"],
        "date_contacted": pd.to_datetime([None, "2024-01-01"]),
        "date_discarded": pd.to_datetime([None, None]),
    })


class UpdateExistingRecordTests(unittest.TestCase):
    def each_page(self):
        for page, function in PAGE_IMPLEMENTATIONS.items():
            with self.subTest(page=page):
                yield function

    def test_updates_the_row_matched_case_insensitively(self):
        for update in self.each_page():
            df, action = update(make_frame(), "ANA sample", {"company": "Initech"})
            self.assertEqual(action, "updated")
            self.assertEqual(df.at[0, "company"], "Initech")
            self.assertEqual(df.at[1, "company"], "Globex")

    def test_name_itself_can_be_renamed(self):
        for update in self.each_page():
            df, _ = update(make_frame(), "ana sample", {"name": "Ana Renamed"})
            self.assertEqual(df.at[0, "name"], "Ana Renamed")

    def test_unknown_name_is_not_found_and_leaves_the_frame_alone(self):
        for update in self.each_page():
            df, action = update(make_frame(), "nobody", {"company": "Initech"})
            self.assertEqual(action, "not_found")
            pd.testing.assert_frame_equal(df, make_frame())

    def test_keys_that_are_not_columns_are_ignored(self):
        for update in self.each_page():
            df, action = update(make_frame(), "ana sample", {"not_a_column": 1})
            self.assertEqual(action, "updated")
            self.assertNotIn("not_a_column", df.columns)

    def test_timestamp_is_stored_in_a_date_column(self):
        for update in self.each_page():
            df, _ = update(make_frame(), "ana sample", {"date_contacted": pd.Timestamp("2025-03-04")})
            self.assertEqual(df.at[0, "date_contacted"], pd.Timestamp("2025-03-04"))
            self.assertTrue(pd.api.types.is_datetime64_any_dtype(df["date_contacted"]))

    def test_clearing_a_date_stores_nat_for_none_and_nat(self):
        for update in self.each_page():
            for cleared in (None, pd.NaT):
                df, _ = update(make_frame(), "bob example", {"date_contacted": cleared})
                self.assertTrue(pd.isna(df.at[1, "date_contacted"]))
                self.assertTrue(pd.api.types.is_datetime64_any_dtype(df["date_contacted"]))

    def test_empty_string_and_iso_string_dates_are_coerced(self):
        for update in self.each_page():
            df, _ = update(make_frame(), "bob example", {"date_contacted": ""})
            self.assertTrue(pd.isna(df.at[1, "date_contacted"]))
            df, _ = update(make_frame(), "bob example", {"date_contacted": "2025-03-04"})
            self.assertEqual(df.at[1, "date_contacted"], pd.Timestamp("2025-03-04"))

    def test_a_row_dict_round_trips_unchanged(self):
        # reachout's Discard button feeds back the selected row plus one new date.
        for update in self.each_page():
            df = make_frame()
            row = df.iloc[1].to_dict()
            row["date_discarded"] = pd.Timestamp("2025-06-01")
            df, _ = update(df, "bob example", row)
            self.assertEqual(df.at[1, "date_contacted"], pd.Timestamp("2024-01-01"))
            self.assertEqual(df.at[1, "date_discarded"], pd.Timestamp("2025-06-01"))

    def test_a_plain_date_object_is_accepted(self):
        # dataentry's old copy raised TypeError here; reachout's coerced it.
        update = _lib.update_existing_record
        df, _ = update(make_frame(), "ana sample", {"date_contacted": datetime.date(2025, 3, 4)})
        self.assertEqual(df.at[0, "date_contacted"], pd.Timestamp("2025-03-04"))

    def test_both_pages_share_the_lib_implementation(self):
        self.assertIs(dataentry.update_existing_record, _lib.update_existing_record)
        self.assertIs(reachout.update_existing_record, _lib.update_existing_record)


if __name__ == "__main__":
    unittest.main()
