"""Regression tests for the quarter date filter (synthetic data only)."""

import os
import sys
import unittest
from datetime import datetime

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import csv_structure_compare as compare  # noqa: E402

DATE_COLUMN = "Created date (UTC)"


def _frame(*stamps: str) -> pd.DataFrame:
    return pd.DataFrame({"id": range(len(stamps)), DATE_COLUMN: list(stamps)})


class QuarterEndFilterTests(unittest.TestCase):
    def test_payments_on_the_last_day_of_the_quarter_are_kept(self):
        # Stripe's "Created date (UTC)" carries a time of day, but the quarter's
        # end_date is midnight of its last day: `<= end_date` dropped every
        # payment made on that day after 00:00.
        start, end = compare.get_last_closed_quarter_dates(datetime(2026, 1, 11))
        df = _frame(
            "2025-09-30 23:59:59",  # previous quarter
            "2025-10-01 00:00:00",  # first instant of Q4
            "2025-12-31 00:00:00",  # last day, midnight
            "2025-12-31 18:45:12",  # last day, after midnight
            "2025-12-31 23:59:59",  # last second of Q4
            "2026-01-01 00:00:00",  # next quarter
        )
        kept = compare.filter_by_date_range(df, DATE_COLUMN, start, end)
        self.assertEqual(
            [str(v) for v in kept[DATE_COLUMN]],
            [
                "2025-10-01 00:00:00",
                "2025-12-31 00:00:00",
                "2025-12-31 18:45:12",
                "2025-12-31 23:59:59",
            ],
        )

    def test_explicit_custom_end_date_includes_the_whole_day(self):
        df = _frame("2024-03-31 15:30:00", "2024-04-01 00:00:00")
        kept = compare.filter_by_date_range(
            df, DATE_COLUMN, datetime(2024, 1, 1), datetime(2024, 3, 31)
        )
        self.assertEqual([str(v) for v in kept[DATE_COLUMN]], ["2024-03-31 15:30:00"])


if __name__ == "__main__":
    unittest.main()
