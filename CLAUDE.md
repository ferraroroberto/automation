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

All seventeen must exit 0 / report `OK`.

Conditional: after changing `google/gmail_drive_automation.py`, also run its offline diagnostics (they import the module and call its `build_sheet_row()` and `filter_new_emails()`) — `cd google && ..\.venv\Scripts\python.exe diag_gmail_drive.py`. Must run from inside `google/`; its sample-config lookup is CWD-relative and resolves wrong from the repo root.

## Where code goes

[`docs/layout.md`](docs/layout.md) lists every top-level folder and subpackage and says where new code and tests go. Tests live under `tests/<domain>/`, mirroring the tool path, never beside the code.

A structure ratchet gates the verification run. `scripts/structure_ratchet.py` is vendored byte-for-byte from project-scaffolding (re-vendor with `/propagate-vendored structure_ratchet`, never edit) and fails when a file passes 1,000 lines, a function 100 lines or a flat directory 40 modules and is new or above its entry in `.structure-ratchet-baseline.json`, or when the root folders do not match `shape = "pipeline-cli"` plus `extra_roots` in `.fleet.toml`. The baselined giants (three files, 27 functions) are recorded at their real size, never excluded: they cannot grow, so new code goes in a new module beside them. Burn the baseline down with `--update-baseline`; only `--allow-additions` adds, and it shows in review. A new tool folder is declared in `extra_roots` and gets a line in `docs/layout.md`. Exit 3 (could not measure) is NOT CONFIRMED and fails the gate. Config is the `[structure_ratchet]` table in `.fleet.toml`.

## Internal architecture

[`docs/architecture.mmd`](docs/architecture.mmd) is a hand-authored Mermaid diagram of this repo's internal structure — the domain folders, shared per-domain helpers (`google/_auth.py`, `notion/utils.py`), and the external services each domain talks to. Update it in the same PR as any material structural change: a new domain folder, a shared helper added/moved, a script relocated. Not auto-generated, not covered by any test suite.
