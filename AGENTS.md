# Automation Project Rules & Standards

This is the concise, general-purpose hub for coding standards and workflow. Use it as the single starting point; deeper topics are linked at the end.

## 🔎 TL;DR Core Rules {#tldr}
- **Name clearly**: files `snake_case.py`, functions/vars `snake_case`, classes `PascalCase`, constants `UPPER_SNAKE_CASE`.
- **Imports order**: stdlib → third‑party → local.
- **Type hints**: annotate functions; use `Optional`, `Union`, `Dict[str, Any]`, `List[T]` when relevant.
- **Logging, not print**: module-level loggers; info/debug with emojis; no persistent file logging by default.
- **Error handling**: fail fast, clear messages, retry only when it helps.
- **Config first**: JSON config; no hardcoding; defaults + validation; secrets via env/.env.
- **SQL style**: descriptive aliases, `COALESCE`, UPPERCASE identifiers, doc complex processes.
- **Security**: never commit secrets; minimal dependencies; pin versions.
- **Plan before code**: write a brief implementation plan and wait for approval.
- **Scope discipline**: implement only what was requested.
- **.venv policy**: never activate; call the interpreter inside `.venv` directly.

## 🐍 Core Standards {#standards}

### Naming & Imports
```python
# Order imports: stdlib, third‑party, local
from pathlib import Path
import requests
from .utils import load_config
```

### Type Hints
- Annotate parameters and returns. Prefer explicit `Optional[T]` instead of bare `None`.

### Logging
```python
logger.debug("📂 Loading configuration file")
logger.info("✅ Configuration loaded successfully")
logger.error(f"❌ Config not found at {config_path}")
```

### Error Handling
- Handle expected errors with helpful messages and context; re‑raise unexpected ones.

### Configuration Rules
- JSON-based configs; no hardcoding of paths/URLs.
- Provide sane defaults; validate on load.
- Use env/.env for secrets; never commit.

## 🗄️ SQL & BigQuery Standards {#sql-standards}
```sql
-- Use descriptive CTE aliases
WITH source_with_flags AS (...)
SELECT COALESCE(a.field, b.field) AS field
FROM PROJECT.DATASET.TABLE AS T
```
- Document complex processes in a sibling `.md`.
- Partition/cluster appropriately for large tables.

## 📝 Documentation {#documentation}
- Use clear docstrings and minimal inline comments for non‑obvious logic.
- README template:
```markdown
# Module Name

## 🚀 Overview
## 📋 Usage
## 🔧 Configuration
## 📊 Output
```

## 📋 Planning Workflow {#planning-workflow}
Before any code changes, provide and get approval for a short plan:
```markdown
## 📋 Implementation Plan

### 🎯 Objective
[What will be accomplished]

### 📁 Files to Modify
- path/to/file1.py — [create/modify/delete] — [what changes]
- path/to/file2.py — [create/modify/delete] — [what changes]

### 🔗 Dependencies
[New packages/imports/external]

### 🧪 Testing Strategy
[How you will validate]

### ⚠️ Risks
[Potential impacts]
```

### Scope Limitation Rule
- No extra features or “nice‑to‑haves” unless explicitly approved.

## 🤖 Virtual Environment Policy {#venv-policy}
Always use the local `.venv` without activation. Call the interpreter directly.

**Linux/macOS:**
```bash
./.venv/bin/python -m pip install -r requirements.txt
./.venv/bin/python -m pytest
```

**Windows (PowerShell):**
```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m pytest
```

Use the same pattern for tools (black, isort, etc.): `./.venv/bin/python -m black .`

## 🖥️ Platform & Shell Assumptions {#platform-shell}
- Default environment: **Windows 10+ with PowerShell**.
- When executing commands with tools, **use PowerShell syntax**, not Unix/bash-only constructs.
- Prefer Windows path formats and quoting (e.g., `".\\venv\\Scripts\\python.exe"`).
- Use PowerShell-style file paths for Git and all CLI commands. Example: `git add .\\automation\\AGENTS.md` (not `git add automation/AGENTS.md`).
- Provide Windows examples first. Unix/macOS equivalents live in `AGENTS_PR.md` and are reference-only.
- Combine with the `.venv` policy above (no activation; call the interpreter directly).

## 📄 GIT commits and push rules
- Do not automatically create commits or push. Wait for explicit instruction from the user. On task completion, ask: "shall I create a commit message, stage, commit and push"?

## 🔒 Security & Dependencies
- Never commit secrets; use env/.env.
- Minimal dependencies; pin versions; keep security patches up to date.

## 🔗 See Also
- [CLI reference](AGENTS_CLI.md#cli-architecture)
- [Structure guide](AGENTS_STRUCTURE.md#reorg-checklist)
- [PR template](AGENTS_PR.md#environment-setup)