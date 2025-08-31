# Terminal Setup: PowerShell 7 in Cursor

## 🎯 Overview
This guide explains how to set up PowerShell 7 as the default terminal in Cursor to enable modern PowerShell features like the `&&` operator and improved command chaining.

## 🔍 Problem
- PowerShell 5.1 (default on Windows) doesn't support `&&` operator
- Commands get "stuck" when using `&&` in PowerShell 5.1
- Multi-line strings cause parser issues with `>>` prompts

## ✅ Solution: PowerShell 7

### Step 1: Install PowerShell 7
```powershell
# Check if winget is available
winget --version

# Install PowerShell 7
winget install Microsoft.PowerShell

# Verify installation
"C:\Program Files\PowerShell\7\pwsh.exe" --version
```

### Step 2: Configure Cursor Settings

#### Method A: Through Settings UI
1. Open Cursor
2. Press `Ctrl + ,` to open Settings
3. Search for: `terminal.integrated.profiles.windows`
4. Click "Edit in settings.json"

#### Method B: Direct JSON Edit
1. Press `Ctrl + Shift + P`
2. Type: `Preferences: Open Settings (JSON)`
3. Press Enter

### Step 3: Update Settings JSON

Replace or update your `terminal.integrated.profiles.windows` section:

```json
{
    "terminal.integrated.profiles.windows": {
        "PowerShell": {
            "path": "C:\\Program Files\\PowerShell\\7\\pwsh.exe",
            "icon": "terminal-powershell"
        },
        "PowerShell 5.1": {
            "source": "PowerShell",
            "icon": "terminal-powershell"
        },
        "Command Prompt": {
            "path": [
                "${env:windir}\\Sysnative\\cmd.exe",
                "${env:windir}\\System32\\cmd.exe"
            ],
            "args": [],
            "icon": "terminal-cmd"
        },
        "Git Bash": {
            "source": "Git Bash"
        },
        "Ubuntu (WSL)": {
            "path": "C:\\WINDOWS\\System32\\wsl.exe",
            "args": [
                "-d",
                "Ubuntu"
            ]
        }
    },
    "terminal.integrated.defaultProfile.windows": "PowerShell"
}
```

## 🧪 Testing

### Verify PowerShell 7 is Active
```powershell
$PSVersionTable.PSVersion
# Should show: 7.x.x instead of 5.1.x.x
```

### Test && Operator
```powershell
Write-Host "First command" && Write-Host "Second command"
# Should execute both commands successfully
```

### Test Multi-line Commands
```powershell
git commit -m "First line" `
           -m "Second line" `
           -m "Third line"
# Should work without getting stuck
```

## 🔧 Troubleshooting

### PowerShell 7 Not Found
```powershell
# Check if PowerShell 7 is installed
Get-Command pwsh -ErrorAction SilentlyContinue

# If not found, install manually
winget install Microsoft.PowerShell
```

### Settings Not Applied
1. Restart Cursor completely
2. Open new terminal (`Ctrl + `` `)
3. Check terminal dropdown shows "PowerShell"

### Fallback to PowerShell 5.1
If needed, you can temporarily switch:
1. Open terminal dropdown (next to +)
2. Select "PowerShell 5.1"

## 📋 Available Terminal Profiles

After setup, you'll have:
- **PowerShell** (PowerShell 7) - Default
- **PowerShell 5.1** - Legacy version
- **Command Prompt** - Windows CMD
- **Git Bash** - Git for Windows bash
- **Ubuntu (WSL)** - WSL Ubuntu

## 🎯 Benefits of PowerShell 7

- ✅ `&&` and `||` operators work
- ✅ Better error handling
- ✅ Improved performance
- ✅ Cross-platform compatibility
- ✅ Modern PowerShell features
- ✅ Better JSON handling
- ✅ Enhanced debugging

## 🔄 Migration Notes

### From PowerShell 5.1 to 7
- Most scripts work without changes
- Some modules may need updates
- Execution policy might need adjustment

### Execution Policy
```powershell
# Check current policy
Get-ExecutionPolicy

# Set if needed (run as admin)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## 📁 File Locations

### Cursor Settings
- **Windows**: `%APPDATA%\Cursor\User\settings.json`
- **macOS**: `~/Library/Application Support/Cursor/User/settings.json`
- **Linux**: `~/.config/Cursor/User/settings.json`

### PowerShell 7 Installation
- **Default Path**: `C:\Program Files\PowerShell\7\pwsh.exe`
- **User Path**: Usually added to PATH automatically

## 🚀 Quick Commands

### Install PowerShell 7
```powershell
winget install Microsoft.PowerShell
```

### Test Installation
```powershell
& "C:\Program Files\PowerShell\7\pwsh.exe" -NoProfile -Command "Write-Host 'PowerShell 7 is working!'"
```

### Test && Operator
```powershell
& "C:\Program Files\PowerShell\7\pwsh.exe" -NoProfile -Command "Write-Host 'First' && Write-Host 'Second'"
```

---

**Last Updated**: 2025-01-30  
**Tested On**: Windows 10/11, Cursor 1.5.7, PowerShell 7.5.2
