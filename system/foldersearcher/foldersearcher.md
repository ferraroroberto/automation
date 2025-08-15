# Folder Searcher

A simple Python application built with Tkinter that allows you to scan folder structures and search for folders containing specific words. This tool is designed to be faster and more focused than Windows' built-in search, as it only searches for folders, not files.

## Features

- **Folder Structure Scanning**: Scan any folder and its subfolders to create a searchable index
- **Fast Folder Search**: Search for folders containing specific words
- **Flexible Search Scope**: Choose between searching the full scanned structure or just the current active Explorer window path
- **Windows Explorer Integration**: Double-click on search results to open folders directly in Windows Explorer
- **Active Explorer Detection**: Automatically detects the most recently active Windows Explorer window
- **Persistent Data**: Saves folder structure between sessions
- **Simple Interface**: Clean, intuitive Tkinter-based GUI
- **Logging**: Built-in logging for debugging and monitoring

## Requirements

- Python 3.6 or higher
- Windows operating system (for Windows Explorer integration)
- Tkinter (usually included with Python)
- pywin32 (for Windows API access)
- psutil (for process management)

## Installation

1. Download the `foldersearcher.py` file
2. Ensure you have Python installed on your system
3. Install required dependencies:
   ```bash
   pip install pywin32 psutil
   ```
   Or install from the main project requirements:
   ```bash
   pip install -r ../../../requirements.txt
   ```

## Usage

### First Time Setup

1. **Run the application**:
   ```bash
   python foldersearcher.py
   ```

2. **Select a root folder**:
   - Click the "Browse" button to select the folder you want to scan
   - The folder path will be saved in the configuration

3. **Scan the folder structure**:
   - Click "Scan Folder Structure" to analyze all folders and subfolders
   - This creates a `folder_structure.txt` file with the folder hierarchy
   - The scan may take a few moments depending on the folder size

### Searching for Folders

1. **Enter a search term** in the "Search word" field
2. **Choose search scope**:
   - **Full Structure**: Search in the entire scanned folder structure (default)
   - **Current Explorer Path**: Search only in the currently active Windows Explorer window
3. **Click "Search"** to find folders containing that word
4. **Results are displayed** in alphabetical order in the list below
5. **Double-click** on any result to open that folder in Windows Explorer

#### Search Scope Options

- **Full Structure Search**: Searches through the entire scanned folder structure. This is the traditional mode that searches all folders that were scanned initially.

- **Current Explorer Path Search**: Searches only within the folder that is currently open in your active Windows Explorer window. This is useful when you want to narrow your search to a specific location you're currently working in.

  - The application automatically detects the most recently active Explorer window
  - If no Explorer window is found, it falls back to full structure search
  - This mode is perfect for quick, focused searches within a specific directory

### Configuration

The application uses two files for configuration:

- `foldersearcher.json`: Contains the root folder path and structure file location
- `folder_structure.txt`: Contains the scanned folder structure (created automatically)

## File Structure

```
foldersearcher.py          # Main application file
foldersearcher.json        # Configuration file
folder_structure.txt       # Scanned folder structure (auto-generated)
README.md                 # This documentation
```

## How It Works

1. **Scanning**: The application walks through all directories in the selected root folder and creates a map of folder names and their relationships
2. **Storage**: The folder structure is saved to a text file for persistence between sessions
3. **Searching**: When you search, the application looks for folders whose names contain your search term (case-insensitive)
4. **Opening**: Double-clicking a result uses Windows' `explorer` command to open the folder

## Benefits Over Windows Search

- **Faster**: No need to scan files, only folder names
- **Focused**: Only shows folders, not files
- **Persistent**: Once scanned, searches are instant
- **Customizable**: You control which folder structure to search

## Logging

The application includes comprehensive logging that shows:
- Application startup and initialization
- Configuration loading/saving
- Folder scanning progress
- Search operations
- Error messages

Logs are displayed in the console and help with debugging if issues occur.

## Troubleshooting

### Common Issues

1. **"No structure file found"**: You need to scan a folder first before searching
2. **"Please select a valid folder"**: Make sure the folder path exists and is accessible
3. **Search returns no results**: Check that your search term is spelled correctly and that folders with that name exist
4. **"No active Explorer window found"**: Make sure you have a Windows Explorer window open and active when using the "Search in current Explorer window path" option
5. **Import errors for win32gui or psutil**: Install the required dependencies with `pip install pywin32 psutil`

### Error Messages

- Check the console output for detailed error messages
- Ensure you have write permissions in the application directory
- Make sure the selected folder is accessible

## Technical Details

- **Language**: Python 3
- **GUI Framework**: Tkinter
- **File Formats**: JSON (configuration), TXT (folder structure)
- **Platform**: Windows (uses Windows Explorer integration)
- **Dependencies**: None (uses only Python standard library)

## License

This project is open source and available under the MIT License. 