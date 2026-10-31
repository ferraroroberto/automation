"""Reachout Manager's two-step delete must actually reach the deletion (Streamlit AppTest, no Excel)."""

import sys
import unittest
from pathlib import Path
from unittest import mock

from streamlit.testing.v1 import AppTest

COMMON = Path(__file__).resolve().parents[4] / "linkedin/profiles_data_extractor/common"
sys.path.insert(0, str(COMMON))

import reachout  # noqa: E402

NAME = "Ann Example"


def _delete_panel(common_path):
    import sys

    import pandas as pd

    sys.path.insert(0, common_path)
    import reachout

    frame = pd.DataFrame({"name": ["Ann Example", "Bob Example"], "url_profile": ["a", "b"]})
    reachout._render_delete_section(frame, "unused.xlsx")


def _app() -> AppTest:
    at = AppTest.from_function(_delete_panel, args=(str(COMMON),))
    at.session_state["reachout_original_name"] = NAME
    at.session_state["reachout_selected_record"] = {"name": NAME}
    at.session_state["reachout_editing"] = True
    return at


class DeleteConfirmationTests(unittest.TestCase):
    def setUp(self):
        saver = mock.patch.object(reachout, "save_to_excel", return_value=True)
        history = mock.patch.object(reachout, "log_history")
        self.save = saver.start()
        self.history = history.start()
        self.addCleanup(saver.stop)
        self.addCleanup(history.stop)

    def test_confirming_deletes_only_the_named_profile(self):
        at = _app().run()
        at.button(key="delete_profile").click().run()
        self.assertEqual(len(at.warning), 1, "the confirmation prompt must stay up after the rerun")

        at.button(key=f"confirm_yes_{NAME}").click().run()

        self.save.assert_called_once()
        saved_frame = self.save.call_args.args[0]
        self.assertEqual(saved_frame["name"].tolist(), ["Bob Example"])
        self.history.assert_called_once()
        self.assertIsNone(at.session_state["reachout_original_name"])

    def test_cancel_removes_the_prompt_and_deletes_nothing(self):
        at = _app().run()
        at.button(key="delete_profile").click().run()
        at.button(key=f"confirm_no_{NAME}").click().run()
        self.assertNotIn(f"confirm_delete_{NAME}", at.session_state)

        at.run()  # AppTest reports the pre-rerun tree after st.rerun(); a plain run shows the settled page
        self.assertEqual(len(at.warning), 0)
        self.save.assert_not_called()


if __name__ == "__main__":
    unittest.main()
