# Python Virtual Environment Manager

## Overview

The `venv_manager.py` script helps you manage your Python environments, both global and virtual. It provides the following features:

- Detect whether you're in a global Python environment or a virtual environment
- Assist with deactivating virtual environments
- List all installed packages in the current environment
- Clean up environments by removing unnecessary packages while preserving essential ones

## Use Cases

- Clean up a bloated global Python installation
- Reset a virtual environment to a minimal state
- Identify which Python environment is currently active
- Learn what packages are installed in your current environment

## How to Use

### Running the Script

Run the script from the command line:

```
python venv_manager.py
```

### Environment Detection

The script will automatically detect and display information about your current Python environment:

```
==================================================
Current Python Environment: [environment name]
Status: [VIRTUAL ENVIRONMENT or GLOBAL PYTHON]
Path: [path to environment]
Python Executable: [path to python executable]
==================================================
```

### Virtual Environment Handling

If you're in a virtual environment, you'll be given options:

1. **Deactivate it before proceeding (y)**: The script will provide instructions to deactivate the VENV
2. **Force continue (f)**: Continue running in the virtual environment despite warnings
3. **Continue normally (n)**: Proceed in the virtual environment without special handling

### Package Management

The script will:

1. Display all installed packages in your current environment
2. Ask if you want to clean the environment by removing unnecessary packages
3. Let you specify which packages to keep (in addition to pip, setuptools, and wheel)
4. Show you what will be uninstalled and ask for confirmation
5. Remove selected packages, leaving you with a cleaner environment

## Features

- Automatic environment detection (virtual or global Python)
- Package listing and visualization
- Safe environment cleaning with essential package preservation
- **Automatic privilege elevation** on Windows for global Python installations
- Multi-platform support
- Detailed environment information

## Tips

- To properly deactivate a virtual environment, you need to run the `deactivate` command in your terminal or command prompt, not within the script.
- If you're having trouble getting out of a virtual environment, try running the script with the global Python interpreter directly:
  ```
  C:\Path\To\GlobalPython\python.exe venv_manager.py
  ```
- The script preserves `pip`, `setuptools`, and `wheel` by default, as these are essential for Python package management.
- Be careful when cleaning your global Python environment, as it might remove packages needed by other applications.

## Troubleshooting

### Permission Errors

If you encounter permission errors like `PermissionError: [WinError 5] Access is denied` when uninstalling packages:

1. **Use Automatic Elevation**: The script can automatically re-launch itself with administrator privileges when needed

2. **Run as Administrator**: Manually restart the script with administrator privileges
   - Right-click on Command Prompt/PowerShell and select "Run as administrator"
   - Navigate to the script location and run it again

3. **Use Virtual Environments**: Instead of modifying the global Python installation, create and use virtual environments:
   ```
   python -m venv .venv
   .venv\Scripts\activate  # Windows
   source .venv/bin/activate  # macOS/Linux
   ```

4. **Manual Package Management**: For stubborn packages, try:
   ```
   pip uninstall --user packagename
   ```

5. **Alternative Global Package Location**: Install packages with the `--user` flag to avoid permission issues:
   ```
   pip install --user packagename
   ```

If the script keeps detecting a virtual environment even after you've deactivated it:

1. Make sure you're using the correct Python executable (use `which python` on Unix or `where python` on Windows to check)
2. Try using the "force continue" option if you're confident about your environment
3. Check if the `VIRTUAL_ENV` environment variable is set in your terminal (use `echo %VIRTUAL_ENV%` on Windows or `echo $VIRTUAL_ENV` on Unix)

## Technical Details

The script detects virtual environments through multiple methods:
- Checking Python's internal environment attributes
- Looking for the `VIRTUAL_ENV` environment variable
- Analyzing the Python executable path for typical VENV indicators

This multi-layered approach ensures accurate environment detection in most scenarios.
