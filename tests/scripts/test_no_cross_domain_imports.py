"""A tool domain may import `system._lib` and the repo-root helpers, never another domain.

The domains (`google/`, `parking/`, `notion/`, ...) are independent scheduled
jobs and tools. A helper two of them need belongs in `system/_lib/`
(`docs/layout.md`); importing it from a sibling domain's package couples their
release behaviour, so a change to one silently changes the other. `html/` is
left out because its name is also the stdlib module `html`.
"""

from __future__ import annotations

import ast
import unittest
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
SKIPPED_DIRS = {"tests", "scripts", "docs", "node_modules", "html"}
SHARED_DOMAIN = "system"


def _domains() -> list[str]:
    return sorted(
        p.name for p in REPO_ROOT.iterdir()
        if p.is_dir() and not p.name.startswith((".", "_")) and p.name not in SKIPPED_DIRS
        and any(p.rglob("*.py"))
    )


def _imported_roots(tree: ast.Module) -> set[str]:
    roots: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            roots.update(alias.name.split(".")[0] for alias in node.names)
        elif isinstance(node, ast.ImportFrom) and node.level == 0 and node.module:
            roots.add(node.module.split(".")[0])
    return roots


class NoCrossDomainImports(unittest.TestCase):
    def test_domains_import_only_themselves_and_the_shared_package(self) -> None:
        domains = _domains()
        self.assertIn("parking", domains)
        offenders = []
        for domain in domains:
            allowed = {domain, SHARED_DOMAIN}
            for path in sorted((REPO_ROOT / domain).rglob("*.py")):
                if ".venv" in path.parts:
                    continue
                tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
                for other in sorted(_imported_roots(tree) & set(domains) - allowed):
                    offenders.append(f"{path.relative_to(REPO_ROOT).as_posix()} imports {other}")
        self.assertEqual(offenders, [], "move the shared helper into system/_lib/ and import it from there")


if __name__ == "__main__":
    unittest.main()
