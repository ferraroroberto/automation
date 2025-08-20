# Apple Contacts Converter

A simple Tkinter GUI application that converts vCard (.vcf) files to Apple-compatible format.

## 🚀 What It Does

- **Select any .vcf file** from your computer using a file browser
- **Automatically converts** it to Apple-compatible vCard 3.0 format
- **Saves output** with `_apple` suffix in the same folder
- **Validates format** for Apple import compatibility
- **No configuration needed** - works out of the box

## 📋 Features

- **Simple GUI**: Easy-to-use Tkinter interface
- **File Browser**: Click to select any .vcf file
- **Automatic Naming**: Output file gets `_apple` suffix automatically
- **Same Folder**: Output is saved in the same location as input
- **Format Validation**: Checks Apple compatibility before saving
- **Real-time Status**: Shows conversion progress and results

## 🍎 Apple Compatibility

The converter ensures:
- vCard 3.0 format (Apple's preferred version)
- Proper TYPE formatting for phone, email, and address fields
- Required fields (FN, VERSION) are present
- CRLF line endings for maximum compatibility
- UTF-8 encoding support

## 🚀 Quick Start

### Windows Users
1. Double-click `run_converter.bat`
2. Or run: `python convert_to_apple.py`

### All Platforms
```bash
python convert_to_apple.py
```

## 📖 How to Use

1. **Launch the Application**
   - Run the script or batch file
   - GUI window will open

2. **Select Input File**
   - Click "Browse" button
   - Choose any .vcf file from your computer
   - Output filename will be shown automatically

3. **Convert**
   - Click "Convert to Apple Format" button
   - Watch the progress in the status area
   - Success message will appear when done

4. **Find Output**
   - Look for `filename_apple.vcf` in the same folder
   - This file is ready for Apple import

## 📁 Files

```
automation/system/apple/
├── convert_to_apple.py      # Main GUI application
├── run_converter.bat        # Windows launcher (double-click to run)
├── convert_to_apple.md      # This documentation
├── contacts.vcf             # Your original file (example)
└── contacts_apple.vcf       # Converted Apple format (output)
```

## 🔧 Requirements

- **Python 3.7+** (includes Tkinter by default)
- **No external packages** - uses only Python standard library
- **Windows, Mac, or Linux** - Tkinter works on all platforms

## 🍎 Importing to Apple

### Supported Platforms
- **iCloud Contacts**: Web interface
- **iPhone/iPad**: Settings > Contacts > Import Contacts
- **Mac**: Contacts app > File > Import
- **Apple Watch**: Syncs automatically

### Import Methods
1. **Email**: Send .vcf file as attachment
2. **AirDrop**: Transfer between Apple devices
3. **iCloud Drive**: Upload and import
4. **Direct Import**: Use device's import function

## 📝 Example

**Input file**: `my_contacts.vcf`
**Output file**: `my_contacts_apple.vcf`

The converter will:
- Read your original vCard file
- Convert to Apple-compatible format
- Save as `my_contacts_apple.vcf` in the same folder
- Validate the format for Apple import

## 🔧 Troubleshooting

### Common Issues

1. **"Python not found"**
   - Install Python 3.7+ from python.org
   - Ensure Python is in your PATH

2. **"No module named tkinter"**
   - Install tkinter: `sudo apt-get install python3-tk` (Linux)
   - Windows/Mac: Tkinter is included by default

3. **Conversion fails**
   - Check file is valid .vcf format
   - Ensure file is not corrupted
   - Check file permissions

4. **Import fails on Apple device**
   - Verify file size is under 10MB
   - Check file format validation passed
   - Try different import method

## 📝 Notes

- **No configuration needed** - works out of the box
- **Preserves all contact data** - no information is lost
- **Safe conversion** - original file is never modified
- **Automatic validation** - ensures Apple compatibility
- **Status logging** - shows detailed progress in the GUI

## 🤝 Support

If you encounter issues:
1. Check the status messages in the GUI
2. Verify your input file is a valid .vcf
3. Ensure Python and Tkinter are properly installed
4. Check file permissions and folder access

The application shows detailed status information in the GUI. Check the status area for any error messages or warnings.
