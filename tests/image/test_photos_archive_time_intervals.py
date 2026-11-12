"""Burst detection diffs capture times in time order, not scan order."""

import sys
import unittest
from pathlib import Path
from unittest import mock

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "image"))

import photos_archive_duplicates  # noqa: E402


def _frame(names):
    return pd.DataFrame({"unified_name": names})


class TimeIntervalTests(unittest.TestCase):
    def test_scan_order_does_not_produce_negative_gaps(self):
        # Two source folders scanned one after the other: 10:00:00, 10:00:30 then 09:00:00.
        df = photos_archive_duplicates.calculate_time_intervals(
            _frame(["20240101-100000.jpg", "20240101-100030.jpg", "20240101-090000.jpg"])
        )
        self.assertEqual(df["time_interval"].dropna().tolist(), [3600.0, 30.0])
        self.assertTrue(pd.isna(df.loc[2, "time_interval"]))

    def test_first_file_of_each_day_has_no_gap(self):
        df = photos_archive_duplicates.calculate_time_intervals(
            _frame(["20240101-100000.jpg", "20240102-100000.jpg"])
        )
        self.assertTrue(df["time_interval"].isna().all())

    def test_repeated_index_labels_map_back_row_for_row(self):
        df = _frame(["20240101-100030.jpg", "20240101-100000.jpg"])
        df.index = [7, 7]
        out = photos_archive_duplicates.calculate_time_intervals(df)
        self.assertEqual(out["time_interval"].tolist()[0], 30.0)
        self.assertTrue(pd.isna(out["time_interval"].tolist()[1]))


class ShortSequenceTests(unittest.TestCase):
    def test_a_spread_out_day_is_not_counted_as_a_burst(self):
        # Hours apart in scan order that interleaves two folders: no short sequence at all.
        df = _frame(["20240101-100000.jpg", "20240101-140000.jpg", "20240101-090000.jpg"])
        with mock.patch.object(photos_archive_duplicates, "get_config", create=True):
            out = photos_archive_duplicates.calculate_short_sequence_counts(df, threshold=60)
        self.assertEqual(out["short_sequence_count"].tolist(), [0, 0, 0])


if __name__ == "__main__":
    unittest.main()
