"""A fixed browser wait or scroll step in the unattended Google flows is a named constant.

The Saturday Photos job and its Gmail half pause for pages that expose no "finished" signal. A bare
``wait_for_timeout(2_500)`` hides whether that figure was measured or guessed, so nobody dares tune it
when Google slows down; a module constant carries its one-clause reason beside the number.
"""

import ast
import tempfile
import unittest
from pathlib import Path

GOOGLE_DIR = Path(__file__).resolve().parents[2] / "tools" / "google"
# Playwright calls whose literal arguments are tunable timings or distances.
TIMED_CALLS = {"wait_for_timeout", "wheel"}


def _bare_literals(path: Path) -> list[str]:
    hits = []
    for node in ast.walk(ast.parse(path.read_text(encoding="utf-8"))):
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute) and node.func.attr in TIMED_CALLS:
            for arg in node.args:
                if isinstance(arg, ast.Constant) and isinstance(arg.value, (int, float)) and arg.value != 0:
                    hits.append(f"{path.name}:{node.lineno} {node.func.attr}({arg.value})")
    return hits


class SettleWaitTests(unittest.TestCase):
    def test_no_google_tool_passes_a_bare_number_to_a_wait_or_wheel(self):
        hits = [hit for path in sorted(GOOGLE_DIR.glob("*.py")) for hit in _bare_literals(path)]
        self.assertEqual(hits, [], "name the number as a constant with its reason: " + "; ".join(hits))

    def test_the_check_sees_a_bare_wait(self):
        with tempfile.TemporaryDirectory() as tmp:
            sample = Path(tmp) / "sample.py"
            sample.write_text("page.wait_for_timeout(2_500)\npage.mouse.wheel(0, 650)\n", encoding="utf-8")
            self.assertEqual(len(_bare_literals(sample)), 2)


if __name__ == "__main__":
    unittest.main()
