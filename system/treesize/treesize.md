# TreeSize - Folder Size Analyzer

A simple GUI application to analyze folder sizes, similar to TreeSize Professional.

## Features

- Select a folder to analyze
- View folder sizes at the current level
- Double-click folders to navigate deeper
- View top 50 largest files in selected folder
- Choose between "Actual Size" and "Space on Disk" metrics
- Refresh button to clear cache and reload
- Open selected folders in system file explorer
- Reveal selected files in explorer/finder
- Size calculations run in background threads
- Modern UI with theme support
- Real-time logging window

## Usage

1. Run the application:
   ```bash
   python treesize.py
   ```

2. Click "Select Folder" to choose a folder to analyze

3. Double-click on any folder to navigate into it

4. Single-click on a folder to see its top 50 largest files

5. Use the radio buttons to switch between "Actual Size" and "Space on Disk" views

6. Right-side buttons:
   - **Refresh**: Clear all caches and reload the current folder
   - **Open Folder**: Open the selected folder in your system file explorer
   - **Open File Location**: Reveal the selected file in explorer/finder

## Size Metrics

- **Actual Size**: The logical size of files (bytes of data)
- **Space on Disk**: The actual disk space used, accounting for:
  - File system cluster size
  - Compressed files
  - Sparse files
  - OneDrive placeholder files

## Requirements

- Python 3.6+
- tkinter (included with Python)
- Windows: Uses native Win32 APIs for accurate disk space calculation
- Cross-platform: Works on Windows, macOS, and Linux

## How it works

- Uses os.walk() to calculate folder sizes recursively
- Displays sizes in human-readable format (KB, MB, GB, etc.)
- Multi-threaded to prevent UI freezing during calculations
- Caches results for better performance
- Automatically sorts folders by size (largest first)
- On Windows: Uses GetCompressedFileSizeW API for accurate disk usage
- Logging window shows calculation progress and any errors
