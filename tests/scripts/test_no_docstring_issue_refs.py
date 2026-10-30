"""Issue-number narratives inside docstrings slip past `comment_ratchet.py`.

`comment_ratchet.py` (project-scaffolding#411) counts `#NNN` only in real
comment tokens; a module/class/function docstring is a STRING token to
`tokenize`, not a COMMENT, so a docstring that narrates "split out by audit
issue #92" or "dedup: audit issue #64" never shows up there (audit #235). This
test closes that gap for Python product code: it walks every file
`comment_ratchet.in_scope()` would count, extracts every docstring with
`ast.get_docstring`, and reuses `comment_ratchet.find_refs()` so "what counts
as a reference" never drifts between the two checks. Any hit fails -- there is
no baseline, because the fix for this audit's findings is to have zero.
"""

from __future__ import annotations

import ast
import sys
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.append(str(REPO_ROOT / "scripts"))
sys.path.append(str(REPO_ROOT))

import comment_ratchet  # noqa: E402


def _docstring_nodes(tree: ast.Module) -> list[tuple[ast.AST, str]]:
    nodes: list[tuple[ast.AST, str]] = []
    for node in ast.walk(tree):
        if isinstance(node, (ast.Module, ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)):
            doc = ast.get_docstring(node, clean=False)
            if doc:
                nodes.append((node, doc))
    return nodes


def _docstring_start_line(node: ast.AST, doc: str) -> int:
    # The docstring is always the first statement's Expr/Constant node; for a
    # Module there is no line 0, so fall back to 1.
    body = getattr(node, "body", None)
    if body:
        return body[0].lineno
    return 1


class NoDocstringIssueRefs(unittest.TestCase):
    def test_product_code_docstrings_carry_no_issue_references(self) -> None:
        root = REPO_ROOT
        cfg, vendored = comment_ratchet.load_config(root)
        if cfg is None:
            self.skipTest("comment_ratchet is not enabled in .fleet.toml")

        violations: list[str] = []
        for path in comment_ratchet.tracked_files(root):
            if Path(path).suffix.lower() != ".py":
                continue
            if not comment_ratchet.in_scope(path, cfg, vendored):
                continue
            text = (root / path).read_text(encoding="utf-8")
            if comment_ratchet.is_generated(text):
                continue
            try:
                tree = ast.parse(text, filename=path)
            except SyntaxError:
                continue
            for node, doc in _docstring_nodes(tree):
                start = _docstring_start_line(node, doc)
                for ref in comment_ratchet.find_refs(doc, start):
                    violations.append(f"{path}:{ref.line}  {ref.text}")

        self.assertEqual(
            violations,
            [],
            "docstring(s) carry an issue-number reference -- reword to the "
            "present-tense why and drop the history:\n" + "\n".join(violations),
        )


if __name__ == "__main__":
    unittest.main()
