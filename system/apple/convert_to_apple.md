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

Run it through the repo `.venv` from this folder:

**Windows PowerShell:**
```powershell
cd system\apple
..\..\.venv\Scripts\python.exe convert_to_apple.py
```

**Unix/Linux/macOS:**
```bash
cd system/apple
../../.venv/bin/python convert_to_apple.py
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
├── unify_vcards.py          # Interactive duplicate-merging tool (see below)
├── _vcard_fields.py         # Shared vCard TYPE-label parsing helper
└── convert_to_apple.md      # This documentation
```

Input and output `.vcf` files are not kept in the repo — you point the app at your own file, and the converted copy is written next to it.

## 🔁 Companion tool: `unify_vcards.py`

`convert_to_apple.py` changes a vCard's *format*; `unify_vcards.py` changes its *contents* — it merges duplicate contacts in a `.vcf` export before you import it. It is an interactive terminal tool (with a Tkinter file picker for choosing the input), not a GUI app.

**What it does**

- Parses every vCard in the file into name, phone, email, address, and organisation fields (`_vcard_fields.py` supplies the shared TYPE-label parsing that `convert_to_apple.py` also uses).
- Groups likely duplicates by fuzzy name similarity (`difflib.SequenceMatcher`) and by shared normalised phone numbers.
- Walks you through each duplicate group in the terminal: pick which record to keep, merge them, edit the resulting name, or skip. Enter accepts the suggested default.
- Re-runs the duplicate scan on the merged result (up to 10 passes) so newly-adjacent duplicates are caught too.
- Writes the unified vCard plus a Markdown report of every decision.

**Running it**

```powershell
cd system\apple

# Pick the input file with a dialog; outputs land next to it
..\..\.venv\Scripts\python.exe unify_vcards.py

# Or pass paths explicitly: input [output] [report]
..\..\.venv\Scripts\python.exe unify_vcards.py contacts.vcf contacts_unified.vcf report.md
```

With no arguments it opens a file dialog. Output paths default to `<input>_unified.vcf` and `<input>_unification_report.md` in the input's folder. `Ctrl+C` exits cleanly without writing.

Typical order: **unify first, then convert** — dedupe the export with `unify_vcards.py`, then run `convert_to_apple.py` on the unified file to produce the Apple-compatible vCard 3.0.

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
