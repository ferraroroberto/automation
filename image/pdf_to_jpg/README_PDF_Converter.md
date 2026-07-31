# PDF to Images Converter

A self-contained Python application that converts PDF files to JPG images. Can be built into an executable file for easy distribution.

## Features

- 🖼️ Convert PDF files to high-quality JPG images
- 📁 GUI folder selection (Windows)
- ⚡ Parallel processing support for faster conversion
- 🗑️ Option to delete source PDF files after conversion
- 📊 Real-time progress tracking and logging
- 🎯 Configurable DPI settings
- 🚀 Self-contained executable (no Python installation required)

## Requirements

### For Development/Running from Source
- Python 3.7+
- PyMuPDF==1.23.8
- tkinter (usually included with Python)

### For Building Executable
- PyInstaller==6.3.0

## Installation

### Option 1: Install Dependencies and Run from Source

1. Install required packages:
```bash
pip install -r requirements.txt
```

2. Run the application:
```bash
python pdf_to_images_converter.py
```

### Option 2: Build Executable (Recommended)

Use the `build_minimal.bat` script to create a clean executable:

1. **Navigate to the project directory:**
   ```cmd
   cd "path\to\automation\image\pdf_to_jpg"
   ```

2. **Run the build script:**
   ```cmd
   build_minimal.bat
   ```

3. **The executable will be created in the `dist` folder as `PDF_to_Images_Converter_Optimized.exe`**

## Build Process Details

The `build_minimal.bat` script:

- ✅ Creates a temporary virtual environment
- ✅ Installs only required dependencies (PyMuPDF==1.23.8, PyInstaller==6.3.0)
- ✅ Builds the executable with console support (for user interaction)
- ✅ Cleans up all temporary files automatically
- ✅ Leaves only the final `.exe` in the `dist` folder

**Why use this approach?**
- Avoids accidental inclusion of unrelated packages from your global environment
- Ensures clean, minimal executable size
- Fully automated and repeatable
- Great for distribution or sharing with others

## Usage

### Running the Application

1. **Launch the application** (either Python script or executable)
2. **Select a folder** containing PDF files:
   - GUI mode: Use the folder selection dialog
   - CLI mode: Enter the folder path manually
3. **Configure settings**:
   - DPI: Image resolution (default: 150)
   - Delete PDFs: Whether to delete source files after conversion
   - Parallel workers: Number of concurrent processes (default: 4)
4. **Confirm and start conversion**

### Output

- Images are saved in the same folder as the source PDFs
- Naming convention: `{pdf_name}_p{page_number:03d}dp{total_pages:03d}.jpg`
- Example: `document_p001dp005.jpg` (page 1 of 5)

## Configuration Options

### DPI Settings
- **150 DPI**: Good quality, smaller file size (default)
- **300 DPI**: High quality, larger file size
- **72 DPI**: Lower quality, smallest file size

### Parallel Processing
- **1 worker**: Sequential processing (slower but more memory efficient)
- **4 workers**: Default parallel processing (good balance)
- **8+ workers**: Maximum speed (requires more memory)

## File Structure

```
pdf_to_jpg/
├── pdf_to_images_converter.py    # Main application
├── requirements.txt              # Python dependencies
├── build_minimal.bat            # Build script for executable
├── dist/                        # Output directory for executable
│   └── PDF_to_Images_Converter_Optimized.exe
└── README_PDF_Converter.md      # This file
```

## Troubleshooting

### Common Issues

1. **"PyMuPDF is required" error**
   - Install PyMuPDF: `pip install PyMuPDF==1.23.8`

2. **Import error with fitz/frontend modules**
   - **Error**: `RuntimeError: Directory 'static/' does not exist` or import conflicts with `fitz`/`frontend`
   - **Cause**: Conflicting packages installed in your environment
   - **Solution**: 
     ```bash
     # Remove conflicting packages
     pip uninstall -y fitz frontend
     
     # Install the correct PyMuPDF
     pip install --upgrade --force-reinstall PyMuPDF==1.23.8
     ```

3. **"No module named 'ast'" error during build**
   - **Cause**: PyInstaller trying to exclude essential Python modules
   - **Solution**: Use `build_minimal.bat` which uses `--console` instead of `--windowed`

4. **"input(): lost sys.stdin" error when running executable**
   - **Cause**: Executable built with `--windowed` flag (no console)
   - **Solution**: Use `build_minimal.bat` which uses `--console` flag

5. **"tkinter not available" warning**
   - Application will fall back to command-line input
   - Install tkinter if needed (usually included with Python)

6. **Permission errors**
   - Run as administrator if converting files in protected folders
   - Ensure write permissions in the target folder

7. **Memory issues with large PDFs**
   - Reduce number of parallel workers
   - Use lower DPI settings

### Build Issues

**Build fails with module errors**
- Use `build_minimal.bat` which creates a clean virtual environment
- Ensures only required dependencies are installed

**Large executable size**
- Normal for self-contained executables (~23MB for this application)
- Includes all dependencies and Python runtime

**Build takes a long time**
- Normal for PyInstaller builds
- Don't interrupt the process

## Technical Details

### Dependencies
- **PyMuPDF==1.23.8**: PDF processing and image conversion
- **tkinter**: GUI folder selection (optional)
- **concurrent.futures**: Parallel processing
- **pathlib**: Path handling (uses standard library version)
- **logging**: Console logging

### Architecture
- Thread-safe logging with locks
- Parallel processing with ThreadPoolExecutor
- Graceful error handling and user interruption
- Memory-efficient processing
- Automatic package conflict resolution during build

## License

This application is part of the [automation](../../README.md) project and is covered by its MIT license.

## Support

For issues or questions, please check the troubleshooting section above or refer to the main project documentation.
