# 🐍 Python Coding Standards & Patterns

This document details the Python coding standards for the Automation project.
Refer to [AGENTS.md](AGENTS.md) for high-level project rules.

## Naming & Imports
```python
# Order imports: stdlib, third‑party, local
from pathlib import Path
import requests
from .utils import load_config

# Constants: UPPER_SNAKE_CASE
MAX_RETRIES = 3

# Classes: PascalCase
class DataProcessor:
    pass

# Functions/Vars: snake_case
def process_data(data_input):
    pass
```

## Type Hints
- Annotate parameters and returns.
- Prefer explicit `Optional[T]` instead of bare `None`.
- Use `Union`, `Dict[str, Any]`, `List[T]` when relevant.

## Logging
**Rule:** Use module-level loggers. No `print()`.
```python
import logging
logger = logging.getLogger(__name__)

logger.debug("📂 Loading configuration file")
logger.info("✅ Configuration loaded successfully")
logger.error(f"❌ Config not found at {config_path}")
```

## Error Handling
- Handle expected errors with helpful messages and context.
- Re‑raise unexpected ones.
- Fail fast.

## Configuration Rules
- JSON-based configs; no hardcoding of paths/URLs.
- Provide sane defaults; validate on load.
- Use env/.env for secrets; never commit secrets.

## 🗄️ SQL & BigQuery Standards
```sql
-- Use descriptive CTE aliases
WITH source_with_flags AS (...)
SELECT COALESCE(a.field, b.field) AS field
FROM PROJECT.DATASET.TABLE AS T
```
- Document complex processes in a sibling `.md`.
- Partition/cluster appropriately for large tables.

## 📝 Documentation Template
```markdown
# Module Name

## 🚀 Overview
## 📋 Usage
## 🔧 Configuration
## 📊 Output
```
