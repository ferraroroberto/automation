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
& .\.venv\Scripts\python.exe -m unittest discover -s tests/system/wifi -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/audio -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/excel/accounting_stripe -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/linkedin/profiles_data_extractor/devtools -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/linkedin/profiles_data_extractor/common -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/linkedin/check_ip -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/system/apple -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/system/quickdeck -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/image -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s tests/video -p "test_*.py"
& .\.venv\Scripts\python.exe -m pytest -q tests/parking
```

All twenty-five must exit 0 / report `OK`.

Conditional: after changing `google/gmail_drive_automation.py`, also run its offline diagnostics (they import the module and call its `build_sheet_row()` and `filter_new_emails()`) — `cd google && ..\.venv\Scripts\python.exe diag_gmail_drive.py`. Must run from inside `google/`; its sample-config lookup is CWD-relative and resolves wrong from the repo root.

## Where code goes

[`docs/layout.md`](docs/layout.md) lists every top-level folder and subpackage and says where new code and tests go. Tests live under `tests/<domain>/`, mirroring the tool path, never beside the code.

A structure ratchet gates the verification run. `scripts/structure_ratchet.py` is vendored byte-for-byte from project-scaffolding (re-vendor with `/propagate-vendored structure_ratchet`, never edit) and fails when a file passes 1,000 lines, a function 100 lines or a flat directory 40 modules and is new or above its entry in `.structure-ratchet-baseline.json`, or when the root folders do not match `shape = "pipeline-cli"` plus `extra_roots` in `.fleet.toml`. The baselined giants (three files, 27 functions) are recorded at their real size, never excluded: they cannot grow, so new code goes in a new module beside them. Burn the baseline down with `--update-baseline`; only `--allow-additions` adds, and it shows in review. A new tool folder is declared in `extra_roots` and gets a line in `docs/layout.md`. Exit 3 (could not measure) is NOT CONFIRMED and fails the gate. Config is the `[structure_ratchet]` table in `.fleet.toml`.

An instruction ratchet gates the same run. `scripts/instruction_ratchet.py` is vendored byte-for-byte from project-scaffolding (re-vendor with `/propagate-vendored instruction_ratchet`, never edit). It lints this file and `AGENTS.md` with the weekly prompt audit's deterministic rules and fails when a rule's hit count rises above `.instruction-ratchet-baseline.json` or this file passes its 200-line cap. The baseline is empty, so any new hit fails by name: reword the line rather than raising it. Exit 3 is NOT CONFIRMED and fails the gate. Config is the `[instruction_ratchet]` table in `.fleet.toml`.

A doc refs ratchet follows it. `scripts/doc_refs_ratchet.py` is vendored byte-for-byte from project-scaffolding (re-vendor with `/propagate-vendored doc_refs_ratchet`, never edit). It reads `README.md`, this file and `docs/**/*.md` and fails when one names a repo path, `python -m` module or relative link that does not exist and is not recorded in `.doc-refs-ratchet-baseline.json`. The baseline is empty: point a stale reference at what exists, or list an intentional pointer in `allow_refs`. No port is declared in `.fleet.toml`, so ports are not checked. Exit 3 is NOT CONFIRMED and fails the gate. Config is the `[doc_refs_ratchet]` table in `.fleet.toml`.

A dead code ratchet follows it. `scripts/dead_code_ratchet.py` is vendored byte-for-byte from project-scaffolding (re-vendor with `/propagate-vendored dead_code_ratchet`, never edit). It fails when a new function, class or method has no reference anywhere in the tracked tree and is not recorded in `.dead-code-ratchet-baseline.json`. The baseline is empty: call or delete a new definition rather than raising it. A hook something outside Python calls goes in `allow` or `entry_decorators` in the `[dead_code_ratchet]` table of `.fleet.toml`. Exit 3 is NOT CONFIRMED and fails the gate.

A static ratchet follows it. `scripts/static_ratchet.py` is vendored byte-for-byte from project-scaffolding (re-vendor with `/propagate-vendored static_ratchet`, never edit). It fails on a finding that `.static-ratchet-baseline.json` does not record: a duplicated block (`jscpd`, run through `npx`), an unused import or variable, or one of the curated `ruff` bug and security codes. The baseline holds today's findings at their real size and can only shrink: fix the finding rather than raising it, and only `--update-baseline --allow-additions` adds an entry. It needs `npx` and the venv's `ruff`; a detector that could not run is exit 3, NOT CONFIRMED, and fails the gate. Config is the `[static_ratchet]` table in `.fleet.toml`. `tests/scripts/` pins that every ratchet declared in `.fleet.toml` runs as a gate stage that fails on exit 3.

A docs ratchet follows it. `scripts/docs_ratchet.py` is vendored byte-for-byte from project-scaffolding (re-vendor with `/propagate-vendored docs_ratchet`, never edit). It fails when the root `README.md` passes 40 KiB or grows past its entry, or when `README.md` or a `docs/**/*.md` file carries more history markers than `.docs-ratchet-baseline.json` records: an issue reference such as `#123`, a dated heading or a changelog heading. The baseline is empty: state the present-tense why in words, and move detail into `docs/<topic>.md`. Exit 3 is NOT CONFIRMED and fails the gate. Config is the `[docs_ratchet]` table in `.fleet.toml`.

A comment ratchet follows it. `scripts/comment_ratchet.py` is vendored byte-for-byte from project-scaffolding (re-vendor with `/propagate-vendored comment_ratchet`, never edit). It counts issue references in the comments of product code (`.py`, `.js`, `.css`, `.ps1`; tests and vendored scripts are left out) and fails when a file's count rises above `.comment-ratchet-baseline.json`. The baseline holds today's references at their real size and can only shrink: a comment states the present-tense why in words, and the history lives in the issue and `git log`. Only `--update-baseline --allow-additions` adds an entry. Exit 3 is NOT CONFIRMED and fails the gate. Config is the `[comment_ratchet]` table in `.fleet.toml`.

A conventions ratchet closes the list. `scripts/conventions_ratchet.py` is vendored byte-for-byte from project-scaffolding (re-vendor with `/propagate-vendored conventions_ratchet`, never edit). It holds two fleet conventions as `ast` rules and fails when a file's count for a rule rises above `.conventions-ratchet-baseline.json`: `no-window` (a `subprocess` spawn passes `creationflags=`, which here means `no_window.NO_WINDOW`) and `no-drive-path` (no hardcoded drive-letter path in a string). The baseline is empty, so a new hit fails by name: fix it, or keep a reasoned `# <rule>-exempt: <why>` marker on the call. The other rules stay off because the repo has no `app/` or `src/` folder (`logging-not-print`), no constant-time compare helper (`secret-compare`) and no git wrapper (`git-wrapper`). Exit 3 is NOT CONFIRMED and fails the gate. Config is the `[conventions_ratchet]` table in `.fleet.toml`.

## Internal architecture

[`docs/architecture.mmd`](docs/architecture.mmd) is a hand-authored Mermaid diagram of this repo's internal structure — the domain folders, shared per-domain helpers (`google/_auth.py`, `notion/utils.py`), and the external services each domain talks to. Update it in the same PR as any material structural change: a new domain folder, a shared helper added/moved, a script relocated. Not auto-generated, not covered by any test suite.
