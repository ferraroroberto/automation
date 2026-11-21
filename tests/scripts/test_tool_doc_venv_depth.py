"""A tool doc's run-from-the-folder interpreter path must climb to the repo root.

`doc_refs_ratchet.py` reads only inline code and relative links, so a command in
a fenced block such as `..\\.venv\\Scripts\\python.exe tool.py` never reaches it.
When a tool folder moves a level deeper the `.bat` launchers get updated and the
docs keep the old depth, so the command resolves to a `.venv` that does not exist.
This test reads every Markdown file under `tools/` and fails when a `..\\` / `../`
chain in front of `.venv` is not exactly as long as the file's folder is deep.
"""

from __future__ import annotations

import re
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
TOOLS = REPO_ROOT / "tools"

# The lookbehind keeps a chain that is the tail of a longer path (`tools/../.venv`) out.
CHAIN = re.compile(r"(?<![\w./\\])((?:\.\.[\\/])+)\.venv")


class ToolDocVenvDepth(unittest.TestCase):
    def test_interpreter_chain_matches_folder_depth(self) -> None:
        wrong: list[str] = []
        for md in sorted(TOOLS.rglob("*.md")):
            depth = len(md.parent.relative_to(REPO_ROOT).parts)
            for lineno, line in enumerate(md.read_text(encoding="utf-8").splitlines(), 1):
                for match in CHAIN.finditer(line):
                    climbs = len(match.group(1)) // 3
                    if climbs != depth:
                        rel = md.relative_to(REPO_ROOT).as_posix()
                        wrong.append(f"{rel}:{lineno}: climbs {climbs}, folder is {depth} deep")
        self.assertEqual(wrong, [], "stale `.venv` depth in tool docs:\n" + "\n".join(wrong))


if __name__ == "__main__":
    unittest.main()
