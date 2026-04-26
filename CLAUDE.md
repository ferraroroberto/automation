# Project Instructions — Python + Streamlit

## Plan mode is the default
Every non-trivial request starts in plan mode. Non-trivial = anything beyond a one-line fix, a typo, or a question I can answer without touching code.

In plan mode:
- Do NOT edit files, run destructive commands, or commit anything
- Investigate the codebase as needed (read files, search, run read-only commands)
- Resolve ambiguity through questions before proposing a plan
- Present the plan only when you're confident it reflects what I actually want
- Stay in plan mode across rejections — if I push back, revise and re-present, don't bail out to execution

Recommended setting in `.claude/settings.json`:
```json
{ "permissions": { "defaultMode": "plan" } }
```

Exit plan mode only after I explicitly approve. Approval transitions straight to execution in the same turn.

## Asking questions
Ask whenever a decision would be expensive to undo or genuinely ambiguous. One sharp question beats three filler ones. Use multi-choice (2-4 options) when the choice space is bounded — much faster for me to answer than prose.

**Always ask before assuming** any of these:
- File or module location for new code
- Data shape or schema
- Page placement (new page vs. section in existing page)
- `st.session_state` key names and scope
- Caching strategy (`@st.cache_data` TTL vs. `@st.cache_resource`)
- Widget `key=` names and input sources
- Data source (upload, local file, DB via secrets)
- Error and empty-state handling
- Whether to add tests, and at what level

**Don't ask about** things you can determine by reading the code, things I've already specified, or process meta-questions like "is the plan ready?" — that's what plan approval is for.

If multiple reasonable approaches exist, present them as options with tradeoffs. Don't pick silently.

## Before editing
- Re-read any file before modifying it. Don't trust memory across long sessions.
- For files >500 LOC, read in chunks; don't assume you've seen the whole file.
- When renaming a symbol, search separately for: direct calls, type references, string literals, dynamic imports, re-exports, and tests.

## Project layout (standing convention)
- `/app/app.py` — Streamlit entry point, `set_page_config` first, horizontal tabs
- `/app/` — one file per tab
- `/src/` — all non-UI code. Never import streamlit here.
- `/tmp/` — gitignored scratch space, contains `.gitkeep`
- `config.json` (gitignored) + `config.json.example` (committed)
- `.env` for secrets, never committed
- `launch_app.bat` at project root
- `README.md` and `requirements.txt` at project root

## Python conventions
- Config in `config.json`, secrets in `.env`. No hardcoded paths or credentials.
- Use the `logging` module, not `print()`. Emojis welcome: ℹ️ ⚠️ ❌ ✅
- snake_case for files/functions, PascalCase for classes, UPPER_CASE for constants
- Type hints on all public functions. Use `Optional[T]`, never bare `None` returns.
- Imports: stdlib → third-party → local
- Pin versions in `requirements.txt`. Use the existing `.venv`.
- Run Python directly: `& .\.venv\Scripts\python.exe ...` (no activation needed)
- Implement only what was asked. No nice-to-haves.

## Streamlit conventions
- `st.set_page_config(layout="wide", page_title="...")` MUST be the first call
- Use `width='stretch'`, never `use_container_width=True` (deprecated)
- All mutable state in `st.session_state`. No module-level globals.
- `@st.cache_data` for DataFrames/files; `@st.cache_resource` for DB/models
- Every widget needs a stable, explicit `key=`
- UI code only in `/app/`. Data logic stays in `/src/`.
- User feedback via `st.error()` / `st.warning()` / `st.success()`, not `st.write()`

## Phased execution for larger work
Multi-file refactors don't go in a single response. Break into phases of ≤5 files each. Complete phase 1, run verification, wait for my approval, then phase 2. Same rule for any task you'd estimate at >30 minutes of work.

## Verification (before declaring a task done)
- Syntax: `& .\.venv\Scripts\python.exe -m py_compile <file>`
- Lint (if configured): `ruff check .`
- Tests (if any exist): `& .\.venv\Scripts\python.exe -m pytest`
- Streamlit boot check for UI changes: `& .\.venv\Scripts\python.exe -m streamlit run app/app.py --server.headless true`
- If no checker exists, say so explicitly. Don't claim "tests pass" when there are no tests.

## Documentation discipline
For feature work and refactors (not trivial fixes):
- Update `README.md` if usage, config, or output changed
- Add `docs/YYYY-MM-DD-short-description.md` with: what was done, files modified, validation run

For one-line fixes and typos: skip the changelog.

## Git
Never auto-commit or push. Never stage files without being asked. When a task is done, ask: "Shall I prepare the commit message?" When asked, provide a ready-to-copy block:

```bash
git add <files>
git commit -m "type: short description

- detail 1
- detail 2"
```

I run it in my own terminal.

## Senior-dev check
Before finishing, ask: "What would a senior, perfectionist dev reject in review?" If the answer points at duplicated state, inconsistent patterns, or broken architecture *within the file you're already editing*, fix it. Don't expand scope to unrelated files.
