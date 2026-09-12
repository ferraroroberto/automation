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
& .\.venv\Scripts\python.exe -m unittest discover -s system/foldersearcher -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest system.test_local_config_hygiene
& .\.venv\Scripts\python.exe -m unittest discover -s system/wifi -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest discover -s video -p "test_*.py"
```

All five must exit 0 / report `OK`.

Conditional: after changing `google/gmail_drive_automation.py`, also run its offline diagnostics — `cd google && ..\.venv\Scripts\python.exe diag_gmail_drive.py`. Must run from inside `google/`; its sample-config lookup is CWD-relative and resolves wrong from the repo root.

## Internal architecture

[`docs/architecture.mmd`](docs/architecture.mmd) is a hand-authored Mermaid diagram of this repo's internal structure — the domain folders, shared per-domain helpers (`google/_auth.py`, `notion/utils.py`), and the external services each domain talks to. Update it in the same PR as any material structural change: a new domain folder, a shared helper added/moved, a script relocated. Not auto-generated, not covered by any test suite.
