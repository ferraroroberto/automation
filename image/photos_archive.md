# Photo Archive Automation

## 🚀 Overview

A comprehensive Python script that organizes photos and videos from a source folder into a destination archive with intelligent duplicate detection, metadata extraction, and automated file management. The script extracts creation dates from EXIF metadata, video metadata, and filename patterns to create standardized naming conventions.

### Key Features
- **Metadata Extraction**: EXIF data from images, embedded metadata from videos, and pattern-based date parsing from filenames
- **Duplicate Detection**: Identifies duplicates by filename patterns and content (SHA256 hashing)
- **Smart Organization**: Creates unified naming scheme `YYYYMMDD-HHMMSS.extension`
- **Interactive & Automated**: Configurable behavior flags for full automation or interactive control
- **Safety Features**: Preserves originals, confirmation prompts, detailed logging

## 📋 Usage

### Interactive Mode (Default)
```powershell
cd E:\automation\automation\image
python photos_archive.py
```

### Automated Mode
Configure `behavior_flags` in `photos_archive.json` to skip prompts:
```json
{
  "behavior_flags": {
    "use_existing_metadata": false,
    "continue_copy": true,
    "auto_delete_copied_files": true,
    "auto_delete_discarded_files": false
  }
}
```

### Prerequisites
- Python 3.7+
- Required packages: `pandas`, `pillow`, `pymediainfo`, `tkinter`

## 🔧 Configuration

Configuration is managed via `photos_archive.json` following AGENTS.md guidelines.

### Required Settings
```json
{
  "source_folder": "E:/fotos/temporal",
  "destination_folder": "E:/fotos/archivo"
}
```

### Processing Thresholds
```json
{
  "processing_thresholds": {
    "short_sequence_seconds": 10,
    "exclude_days_over_files": 0,
    "exclude_short_sequences_over": 0
  }
}
```

### Behavior Control
```json
{
  "behavior_flags": {
    "skip_destination_folder_scan": false,
    "auto_delete_copied_files": false,
    "auto_delete_discarded_files": false,
    "exclude_creation_modified_criteria": false,
    "use_existing_metadata": false,
    "recalculate_logic": false,
    "continue_copy": false
  }
}
```

### Supported File Types
```json
{
  "file_extensions": {
    "image": [".png", ".jpg", ".jpeg", ".tiff", ".bmp", ".gif"],
    "video": [".mp4", ".mov", ".avi", ".mkv", ".wmv", ".flv"]
  }
}
```

### Date Parsing Patterns
The script recognizes various filename patterns for date extraction:
- `ScreenRecord_2024-01-15-14-30-45`
- `20240115_143045`
- `img_20240115143045`
- `IMG_20240115_143045`
- `VID-20240115-WA001`
- `WP_20240115_001`
- `IMG-20240115-WA001`
- `2024-01-15 14.30.45`
- `Collage 2024-01-15 14_30_45`
- `Screenshot_2024-01-15-14-30-45`

## 📊 Output

### File Organization
Files are renamed using unified format: `YYYYMMDD-HHMMSS.extension`

**Example transformations:**
- `IMG_20240115_143045.jpg` → `20240115-143045.jpg`
- `ScreenRecord_2024-01-15-14-30-45.mp4` → `20240115-143045.mp4`

### Duplicate Handling
- Files with same unified name are marked as duplicates
- SHA256 content hashing available for exact duplicate detection
- Keeps largest file or allows manual selection
- Adds suffixes for non-duplicate similar files: `20240115-143045-001.jpg`

### Metadata Files
Generates Excel files with detailed processing information:
- `metadata_YYYYMMDD-HHMM.xlsx` - Complete file inventory
- Includes columns: file_path, creation_date, criteria, duplicate status, etc.

### Log Files
- `photo_processing_YYYYMMDD-HHMM.log` - Detailed execution log
- Progress reporting every 1000 files processed
- Error tracking and warnings

## 🔄 Processing Workflow

1. **Configuration Load**: Validate JSON config and folder paths
2. **Metadata Collection**: Scan source folder, extract dates from multiple sources
3. **Duplicate Analysis**: Identify filename and content duplicates
4. **Exclusion Rules**: Apply user-defined filters (file counts, date criteria)
5. **File Copy**: Rename and copy qualifying files to destination
6. **Cleanup**: Optionally delete source files and discarded duplicates
7. **Report Generation**: Create Excel metadata and log files

## ⚠️ Safety Features

- **Original Preservation**: Never modifies source files directly
- **Confirmation Prompts**: Interactive approval for destructive operations
- **Path Validation**: Checks folder existence before processing
- **Error Recovery**: Continues processing despite individual file errors
- **Detailed Logging**: Complete audit trail of all operations

## 🛠️ Customization

### Adding New File Types
Update `file_extensions` in config:
```json
{
  "file_extensions": {
    "image": [".png", ".jpg", ".jpeg", ".tiff", ".bmp", ".gif", ".webp"],
    "video": [".mp4", ".mov", ".avi", ".mkv", ".wmv", ".flv", ".m4v"]
  }
}
```

### Custom Date Patterns
Add regex patterns to `date_parsing.filename_patterns` array using standard Python regex syntax.

### Automation Scripts
Create batch files or scheduled tasks using the behavior flags to run unattended backups.