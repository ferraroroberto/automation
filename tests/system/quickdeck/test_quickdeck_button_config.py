"""Slot and field rules of the QuickDeck editor (pure functions, no GUI)."""

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "system/quickdeck"))

from quickdeck_button_config import apply_button_values, clear_slot, join_command, split_command  # noqa: E402


def _values(**overrides):
    values = {"label": "Run", "emoji": "R", "hint": "", "type": "shell", "action": "echo hi"}
    values.update(overrides)
    return values


class ClearSlotTests(unittest.TestCase):
    def test_clearing_a_slot_keeps_later_buttons_in_place(self):
        buttons = [{"label": "A"}, {"label": "B"}, {"label": "C"}]
        self.assertTrue(clear_slot(buttons, 0))
        self.assertEqual(buttons, [{}, {"label": "B"}, {"label": "C"}])

    def test_an_empty_or_missing_slot_reports_nothing_cleared(self):
        buttons = [{}, {"label": "B"}]
        self.assertFalse(clear_slot(buttons, 0))
        self.assertFalse(clear_slot(buttons, 5))
        self.assertEqual(buttons, [{}, {"label": "B"}])


class SplitCommandTests(unittest.TestCase):
    def test_quoted_path_with_spaces_stays_one_argument(self):
        self.assertEqual(
            split_command('python "C:\\Program Files\\tool.py" --flag'),
            ["python", "C:\\Program Files\\tool.py", "--flag"],
        )

    def test_backslashes_in_unquoted_paths_survive(self):
        self.assertEqual(split_command("C:\\tools\\run.bat a"), ["C:\\tools\\run.bat", "a"])

    def test_join_then_split_round_trips_an_argument_with_spaces(self):
        argv = ["python", "C:\\Program Files\\tool.py", "x"]
        self.assertEqual(split_command(join_command(argv[0], argv[1:])), argv)


class ApplyButtonValuesTests(unittest.TestCase):
    def test_saving_a_shell_button_keeps_its_cwd(self):
        existing = {"label": "Old", "type": "shell", "cmd": "make", "cwd": "E:/work"}
        saved = apply_button_values(existing, _values(action="make test"), 1)
        self.assertEqual(saved["cwd"], "E:/work")
        self.assertEqual((saved["cmd"], saved["args"]), ("make", ["test"]))

    def test_changing_type_drops_the_old_types_keys(self):
        existing = {"label": "Old", "type": "shell", "cmd": "make", "args": ["x"], "cwd": "E:/work"}
        saved = apply_button_values(existing, _values(type="copy", action="hello"), 1)
        self.assertEqual(saved["text"], "hello")
        for key in ("cmd", "args", "cwd"):
            self.assertNotIn(key, saved)

    def test_a_non_http_url_is_rejected_not_replaced(self):
        with self.assertRaises(ValueError):
            apply_button_values({}, _values(type="url", action="example.test"), 1)

    def test_an_http_url_is_saved_as_typed(self):
        saved = apply_button_values({}, _values(type="url", action="https://example.test/a"), 1)
        self.assertEqual(saved["url"], "https://example.test/a")

    def test_blank_label_falls_back_to_the_slot_number(self):
        self.assertEqual(apply_button_values({}, _values(label=" "), 3)["label"], "Button 3")

    def test_a_single_word_command_drops_stale_args(self):
        existing = {"type": "shell", "cmd": "make", "args": ["old"]}
        saved = apply_button_values(existing, _values(action="make"), 1)
        self.assertNotIn("args", saved)


if __name__ == "__main__":
    unittest.main()
