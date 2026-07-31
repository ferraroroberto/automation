# Project Instructions

Canonical instructions for AI coding agents working in this repository. Claude Code reads this file directly as project memory. Other agents (Cursor, Codex, etc.) reach it via the one-line `AGENTS.md` pointer.

## Streamlit conventions
*Apply only if this project uses Streamlit.*

- `st.set_page_config(layout="wide", page_title="...")` MUST be the first Streamlit call.
- Use `width="stretch"` (and `width="content"` where appropriate) in new and modified code. **Never** introduce new `use_container_width=True` — it is deprecated. When you touch existing code that uses `use_container_width`, migrate it.
- All mutable state in `st.session_state`. No module-level globals.
- `@st.cache_data` for DataFrames/files; `@st.cache_resource` for DB clients/models.
- Every widget needs a stable, explicit `key=`.
- UI code only in the UI directory (e.g. `app/`). Data logic stays in the non-UI package (e.g. `src/`). Never import `streamlit` from non-UI code.
- User feedback via `st.error()` / `st.warning()` / `st.success()`, not `st.write()`.
- **App layout:** main file (e.g. `app.py`) handles only page config, shared state, sidebar, and tab/radio routing. Each tab/mode lives in its own file exposing a `main(...)` (or `render_*`) function. Default to `st.tabs()`; use a sidebar radio only when asked.

## This repository
Personal automation monorepo: independent Python scripts and small tools across multiple domains (audio, video, image, google, notion, linkedin, system, etc.) on Windows + PowerShell.
See `README.md` for setup, layout, and usage.

## Verification (before declaring a task done)

Run through this repo's own `.venv` by path — a bare `python`/`py` is not reliably on PATH on this machine:

```powershell
& .\.venv\Scripts\python.exe -m compileall -q -x '\.venv' .
& .\.venv\Scripts\python.exe -m unittest discover -s system/foldersearcher -p "test_*.py"
& .\.venv\Scripts\python.exe -m unittest system.test_local_config_hygiene
```

All three must exit 0 / report `OK`. Report failures with the actual output — never claim "tests pass" without having run them. One canonical entry point running the same sequence: `powershell -File scripts\verify-before-ship.ps1`.

Conditional: after changing `google/gmail_drive_automation.py`, also run its offline diagnostics — `cd google && ..\.venv\Scripts\python.exe diag_gmail_drive.py` (must run from inside `google/`; its sample-config lookup is CWD-relative and resolves wrong from the repo root).

## Internal architecture

[`docs/architecture.mmd`](docs/architecture.mmd) is a hand-authored Mermaid diagram of this repo's own internal structure (the domain folders, shared per-domain helpers like `google/_auth.py` and `notion/utils.py`, and the external services each domain talks to). Update it in the same PR as any material structural change (a new domain folder, a shared helper added/moved, a script relocated) — same anti-staleness contract as this repo's `.fleet.toml` `description` field. It is not auto-generated and not covered by any test suite.
