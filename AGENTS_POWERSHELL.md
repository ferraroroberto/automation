# 🖥️ PowerShell & Shell Standards

This document details the shell usage standards for the Automation project.
**Primary Environment:** Windows 10+ with PowerShell 7+.

## ⚡ Core Rules
- **Use PowerShell 7+** for `&&` operator and command chaining.
- **Avoid PowerShell 5.1** (legacy) if possible.
- **Path Separators:** Prefer Windows backslashes `\` and quoting (e.g., `".\\venv\\Scripts\\python.exe"`).
- **Git:** Use PowerShell-style paths: `git add .\\automation\\AGENTS.md`.

## 🔄 Command Patterns

### Chaining (PowerShell 7+)
```powershell
cd "E:\automation\project" && git add . && git commit -m "Update"
```

### Multi-line Commands
```powershell
# Backtick continuation
git commit -m "Add feature" `
           -m "- Core functionality"

# Here-string (for complex messages)
git commit -m @"
Add logging system
- JSON output
- Log rotation
"@
```

### Error Handling
```powershell
git add .; if ($?) { git commit -m "Success" } else { Write-Host "❌ Failed" }
```

## 🛠️ Common Operations

### File Operations
```powershell
New-Item -ItemType File -Path ".\new_script.py" -Force
Copy-Item ".\source\*" ".\destination\" -Recurse
Remove-Item ".\temp\*" -Recurse -Force
```

### Python Execution
Always call the interpreter directly (no activation):
```powershell
# Run script
& ".\.venv\Scripts\python.exe" path\to\script.py

# Run module
& ".\.venv\Scripts\python.exe" -m pytest
```

### Troubleshooting
```powershell
$PSVersionTable.PSVersion
Get-Command python -ErrorAction SilentlyContinue
Get-ChildItem Env: | Where-Object {$_.Name -like "*PYTHON*"}
```
