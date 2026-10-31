# Project Instructions

Canonical agent instructions for this repo; `AGENTS.md` points here.

## This repository

Personal automation monorepo on Windows + PowerShell: independent Python scripts and small tools across domains (`audio`, `video`, `image`, `google`, `notion`, `linkedin`, `system`, …).

## Streamlit

This repo's only Streamlit surface is `linkedin/profiles_data_extractor/common/`. The conventions are owned by `project-scaffolding`'s `CLAUDE.md` § "Streamlit conventions" — read it before touching that code, and never fork a local copy (the copy formerly kept here had already gone stale on tabs-vs-multipage nav). Two that bite existing code here: never introduce new `use_container_width=True` (deprecated — migrate existing uses when you touch them), and never import `streamlit` from non-UI code.

## Verification (before declaring a task done)

One canonical entry point runs the whole sequence: `powershell -File scripts\verify-before-ship.ps1`. It runs, in order:

```powershell
& .\.venv\Scripts\python.exe -m compileall -q -x '\.venv' .
& .\.venv\Scripts\python.exe scripts\structure_ratchet.py
& .\.venv\Scripts\python.exe scripts\instruction_ratchet.py
& .\.venv\Scripts\python.exe scripts\doc_refs_ratchet.py
& .\.venv\Scripts\python.exe scripts\dead_code_ratchet.py
& .\.venv\Scripts\python.exe scripts\static_ratchet.py
& .\.venv\Scripts\python.exe scripts\docs_ratchet.py
& .\.venv\Scripts\python.exe scripts\comment_ratchet.py
& .\.venv\Scripts\python.exe scripts\conventions_ratchet.py
& .\.venv\Scripts\python.exe -m unittest discover -s tests/scripts -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/system/foldersearcher -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/system -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/google -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/notion/articles_sync -p "test*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/notion -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/system/wifi -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/audio -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/excel/accounting_stripe -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/linkedin/profiles_data_extractor/devtools -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/linkedin/profiles_data_extractor/common -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/linkedin/check_ip -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/linkedin/open_profiles -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/system/apple -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/system/quickdeck -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/image -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/video -p "test_*.py"
& .\.venv\Scripts\python.exe -m pytest -q tests/parking
```

All twenty-seven must exit 0 / report `OK`.

Conditional: after changing `google/gmail_drive_automation.py`, also run its offline diagnostics (they import the module and call its `build_sheet_row()` and `filter_new_emails()`) — `cd google && ..\.venv\Scripts\python.exe diag_gmail_drive.py`. Must run from inside `google/`; its sample-config lookup is CWD-relative and resolves wrong from the repo root.

## Where code goes

[`docs/layout.md`](docs/layout.md) lists every top-level folder and subpackage and says where new code and tests go. Tests live under `tests/<domain>/`, mirroring the tool path, never beside the code.

Eight ratchets gate the verification run, in the order listed above. Each `scripts/<name>_ratchet.py` is vendored byte-for-byte from project-scaffolding (re-vendor with `/propagate-vendored <name>_ratchet`, never edit) and fails when a measure rises above its entry in `.<name>-ratchet-baseline.json`. Every baseline is empty, so a new hit fails by name: fix it rather than raising the baseline. Shrink one with `--update-baseline`; only `--allow-additions` adds, and it shows in review. Exit 3 (could not measure) is NOT CONFIRMED and fails the gate. Config is the `[<name>_ratchet]` table in `.fleet.toml`.

- **Structure:** fails when a file passes 1,000 lines, a function 100 lines or a flat directory 40 modules and is new or above its baseline entry, or when the root folders do not match `shape = "pipeline-cli"` plus `extra_roots`. A file near 1,000 lines is split by stage into a module beside it (`notion/articles_sync/` keeps its change detection and property mapping that way), never recorded. A new tool folder is declared in `extra_roots` and gets a line in `docs/layout.md`.
- **Instruction:** lints this file and `AGENTS.md` with the weekly prompt audit's deterministic rules; fails when a rule's hit count rises or this file passes its 200-line cap. Reword the line.
- **Doc refs:** reads `README.md`, this file and `docs/**/*.md`; fails when one names a repo path, `python -m` module or relative link that does not exist. Point a stale reference at what exists, or list an intentional pointer in `allow_refs`. No port is declared in `.fleet.toml`, so ports are not checked.
- **Dead code:** fails when a new function, class or method has no reference anywhere in the tracked tree. Call or delete it; a hook something outside Python calls goes in `allow` or `entry_decorators`.
- **Static:** fails on a duplicated block (`jscpd`, run through `npx`), an unused import or variable, or a curated `ruff` bug or security code. A legitimate exception takes an inline `# noqa: <code> -- <why>`. Needs `npx` and the venv's `ruff`; a detector that could not run is exit 3. `tests/scripts/` pins that every ratchet declared in `.fleet.toml` runs as a gate stage that fails on exit 3.
- **Docs:** fails when the root `README.md` passes 40 KiB or grows past its entry, or when any Markdown file under the root, `docs/` or a tool domain carries more history markers than recorded: an issue reference such as `#123`, a dated heading or a changelog heading (`[docs_ratchet].include` in `.fleet.toml` widens the vendored default from `README.md`/`docs/` to every `extra_roots` folder, since per-tool READMEs carry the same present-tense rule). State the present-tense why in words, and move detail into `docs/<topic>.md`.
- **Comment:** counts issue references in the comments of product code (`.py`, `.js`, `.css`, `.ps1`; tests and vendored scripts are left out). A comment states the present-tense why; the history lives in the issue and `git log`.
- **Conventions:** two `ast` rules, `no-window` (a `subprocess` spawn passes `creationflags=`, which here means `no_window.NO_WINDOW`) and `no-drive-path` (no hardcoded drive-letter path in a string). A legitimate exception keeps a reasoned `# <rule>-exempt: <why>` marker on the call. The other rules stay off: no `app/` or `src/` folder (`logging-not-print`), no constant-time compare helper (`secret-compare`), no git wrapper (`git-wrapper`).

## Internal architecture

[`docs/architecture.mmd`](docs/architecture.mmd) is a hand-authored Mermaid diagram of this repo's internal structure — the domain folders, shared per-domain helpers (`google/_auth.py`, `notion/utils.py`), and the external services each domain talks to. Update it in the same PR as any material structural change: a new domain folder, a shared helper added/moved, a script relocated. Not auto-generated, not covered by any test suite.
