# 🚀 Automation Tools Collection

A comprehensive collection of Python automation tools for audio processing, image manipulation, video recording, email management, Notion integration, and system utilities.

## 🎯 Overview

This repository contains a diverse set of automation tools designed to streamline common digital tasks. Each module is built with a focus on reliability, user experience, and cross-platform compatibility. The tools follow consistent coding standards with proper error handling, logging, and configuration management.

## 🏗️ Project Structure

```
automation/
├── 📁 audio/           # Audio recording, transcription, and conversion tools
├── 📁 excel/           # Excel automation (e.g. Stripe accounting)
├── 📁 google/          # Gmail, Drive, and Google Photos automation
├── 📁 html/            # HTML utilities (e.g. countdown timer)
├── 📁 image/           # Image processing, formatting, and Instagram tools
├── 📁 linkedin/        # LinkedIn reverse-image search, profile opening, and profile data extraction
├── 📁 notion/          # Notion API integration and database management
├── 📁 smart_life/      # Smart Life / IoT device automation
├── 📁 system/          # System utilities and virtual environment management
├── 📁 text/            # Text processing and PDF conversion tools
└── 📁 video/           # Screen recording and video processing tools
```

Configuration uses a root `.env` file (see [Configuration](#️-configuration)); copy from `.env.sample` if present.

### 📚 Agent Instructions

- **`CLAUDE.md`** — canonical instruction set for any AI coding agent working in this repo.
- **`AGENTS.md`** — one-line pointer to `CLAUDE.md` for non-Claude agents (Cursor, Codex, etc.).

The master template lives in [`project-scaffolding/docs/agents/`](../project-scaffolding/docs/agents/) and is propagated to every sibling repo.

## 🔧 Core Modules

### 🎵 Audio Processing (`audio/`)

- **`audio_extractor_core.py`** / **`audio_extractor_gui.py`** - Audio extraction and processing
- **`convert_ogg_mp3.py`** - Audio format conversion (OGG to MP3)
- **`transcript_collate.py`** - Collate and manage transcripts
- **`audio_normalize.py`** - Audio normalization (EBU R128 loudness)
- **`audio_trim.py`** - Interactive audio trimming

Full per-script reference: [`audio/README.md`](audio/README.md).

### 🖼️ Image Processing (`image/`)

**Illustration & Social Media Tools**
- **`illustrations_formatter.py`** - Format illustrations for Instagram or fixed dimensions
- **`illustrations_formatter_gui.py`** - GUI for illustration formatting
- **`photos_archive.py`** - Photo organization and metadata management
- **`image_resizer.py`** - Batch image resizing and optimization

**PDF & Document Processing**
- **`carrousel_pdf_to_jpg.py`** - Convert PDF pages to JPG images
- **`png_to_jpg_converter.py`** - PNG to JPG format conversion

**Utilities**
- **`screenshot.py`** - Automated screenshot capture
- **`collage_image.py`** - Create image collages from multiple files
- **`gif_unpacker.py`** - Extract frames from GIFs
- **`transparency_variants.py`** - Generate transparency variants
- **`heic_to_jpg_converter.py`** - Convert HEIC/HEIF photos to JPG
- **`image_to_pdf_collate.py`** - Collate a folder of images into one PDF
- **`illustrations_check.py`** - Check illustration source/export pairs for orphans
- **`illustrations_subscribers.py`** - Distribute illustrations to subscribers (optional steganographic watermark)
- **`imgur_check_rate.py`** - Report Imgur API rate-limit/quota status
- **`files_list.py`** - List a folder's files with modification dates to JSON
- **`pdf_to_jpg/`** - PDF to JPG conversion (see `README_PDF_Converter.md`)

Reference for the previously-undocumented tools above: [`image/README.md`](image/README.md).

### 📹 Video Processing (`video/`)

**Screen Recording & Capture**
- **`screen_recorder.py`** - Multi-monitor screen recording with mouse overlay — see [`video/screen_recorder.md`](video/screen_recorder.md)
- **`video_trim.py`** - Video trimming and editing utilities
- **`video_concatenator.py`** - Combine multiple video files

**Download & Processing**
- **`video_download.py`** - Unified video downloader (YouTube, HLS/M3U8, direct URL) with GUI
- **`video_reencoder.py`** - Video format conversion and re-encoding

**GPU Maintenance**
- **`gpu_recovery.py`** - Diagnose & recover an NVIDIA GPU stuck "disabled" (Device Manager Code 22); restores NVENC acceleration for the tools above (see `gpu_recovery.md`)

### 📚 Notion Integration (`notion/`)

**Database Management**
- **`build_newsletter.py`** - Newsletter builder from Notion articles — see [`notion/build_newsletter.md`](notion/build_newsletter.md)
- **`notion_databases_dump.py`** - Export Notion databases to Excel
- **`notion_databases_query.py`** - List every database the integration can see and export name/id/url to Excel — see [`notion/notion_databases_query.md`](notion/notion_databases_query.md)
- **`notion_databases_clean.py`** - Apply per-column extract/keep/rename/reorder rules to dumped databases — see [`notion/notion_databases_clean.md`](notion/notion_databases_clean.md)
- **`notion_databases_add_editorial.py`** - Add editorial calendar rows (one day per date) to a Notion database — see [`notion/notion_databases_add_editorial.md`](notion/notion_databases_add_editorial.md)
- **`notion_databases_editorial.py`** - Local Excel-to-Excel refresh of the editorial-calendar workbook (no Notion API calls) — see [`notion/notion_databases_editorial.md`](notion/notion_databases_editorial.md)
- **`notion_excel_sync.py`** - Push rows from an Excel workbook into a Notion database as new pages — see [`notion/notion_excel_sync.md`](notion/notion_excel_sync.md)
- **`sample_illustrations.py`** - Sample illustrations from a Notion database — see [`notion/sample_illustrations.md`](notion/sample_illustrations.md)
- **`articles_sync/`** - Notion articles sync and incremental sync — see [`notion/articles_sync/notion_articles_sync_readme.md`](notion/articles_sync/notion_articles_sync_readme.md)

**Content Processing**
- **`normalize_names.py`** - Name normalization and standardization — see [`notion/normalize_names.md`](notion/normalize_names.md)
- **`normalize_url.py`** - Strip tracking params from URLs in a Notion database — see [`notion/normalize_url.md`](notion/normalize_url.md)
- **`todoist_migration.py`** - Transform a folder of Todoist CSV exports into a consolidated Excel workbook — see [`notion/todoist_migration.md`](notion/todoist_migration.md)
- **`journal_automation.py`** - Automated journal entry creation — see [`notion/journal_automation.md`](notion/journal_automation.md)

### 💻 System Utilities (`system/`)

**Environment & Input**
- **`venv_manager.py`** - Virtual environment management and monitoring
- **`mouse_mover.py`** - Automated mouse movement and clicking
- **`keycaster.py`** - Keyboard input automation and monitoring
- **`quickdeck/`** - Quick Deck / Stream Deck integration
- **`textexpander/`** - Text expander and prompt templates
- **`wifi/`** - Wi‑Fi connection scripts, saved‑password export, BAT generator GUI, and network device scanner — see [`system/wifi/README.md`](system/wifi/README.md)
- **`apple/`** - Apple-related conversion and utilities

**File & Document Operations**
- **`unzip_with_password.py`** - Password-protected archive extraction
- **`rename_files.py`** - Batch file renaming utilities
- **`copy_git_project.py`** - Git project copying and setup
- **`base64_encode_decode.py`** - Base64 encode/decode utility
- **`word_to_markdown.py`** - Word to Markdown conversion
- **`list_files_to_xls.py`** - List a folder's files to an Excel workbook
- **`background.py`** - Set desktop/taskbar color theme (black or light grey)
- **`markdown_preview.py`** - Tray-resident GitHub-style Markdown previewer (Edge WebView2) with light/dark toggle and live reload — see [`system/markdown_preview.md`](system/markdown_preview.md)
- **`foldersearcher/`**, **`treesize/`** - Folder search and size utilities

Reference for the previously-undocumented `system/` scripts (background, keycaster, unzip_with_password, rename_files, copy_git_project, word_to_markdown, list_files_to_xls): [`system/README.md`](system/README.md).

### 📝 Text Processing (`text/`)

**Document Conversion**
- **`convert_pdf_to_txt.py`** - PDF to text extraction
- **`clean_sensitive_data.py`** - Sensitive data removal and sanitization

Full per-script reference: [`text/README.md`](text/README.md).

### 🔗 LinkedIn Tools (`linkedin/`)

**Reverse Image Search & Profile Automation**
- **`check_ip/`** - Reverse-image search (Google Lens via SerpAPI) that uploads images to Imgur and counts where they appear across LinkedIn and other social platforms — see [`linkedin/check_ip/README.md`](linkedin/check_ip/README.md)
- **`open_profiles/`** - Opens a filtered set of LinkedIn profile/company activity pages from an Excel export — see [`linkedin/open_profiles/README.md`](linkedin/open_profiles/README.md)
- **`profiles_data_extractor/`** - DevTools-based profile data extraction, Streamlit dashboard, Excel formatting, and data entry tools

### ☁️ Google Integration (`google/`)

**Gmail, Drive & Photos**
- **`gmail_drive_automation.py`** - Gmail and Google Drive automation — see [`google/gmail_drive_automation.md`](google/gmail_drive_automation.md)
- **`weekly_photo_automation.py`** - Weekly photo automation (e.g. Google Photos) — see [`google/weekly_photo_automation.md`](google/weekly_photo_automation.md)
- **`setup_helper_script.py`** / **`diag_all_scopes.py`** / **`diag_gmail_drive.py`** - First-time config wizard and OAuth-scope / offline diagnostic helpers — see [`google/README.md`](google/README.md)

### 📊 Excel Automation (`excel/`)

- **`accounting_stripe/`** - Stripe accounting and Excel workflows

### 🌐 HTML Utilities (`html/`)

- **`countdown_timer.html`** - Countdown timer and similar utilities

### 🏠 Smart Life / IoT (`smart_life/`)

- Device automation (e.g. **`despacho_switch.bat`**) and configuration via `devices.sample.json`

## 🚀 Quick Start

### **Prerequisites**
- Python 3.8+ installed
- FFmpeg for audio/video processing
- Windows (for Outlook automation tools)
- Git for repository management

### **Environment Setup**

**Windows PowerShell:**
```powershell
# 1. Clone the repository
git clone https://github.com/ferraroroberto/automation.git
cd automation

# 2. Create virtual environment
python -m venv .venv

# 3. Activate virtual environment
.venv\Scripts\Activate.ps1

# 4. Install dependencies
pip install -r requirements.txt
```

**Unix/Linux/macOS:**
```bash
# 1. Clone the repository
git clone https://github.com/ferraroroberto/automation.git
cd automation

# 2. Create virtual environment
python3 -m venv .venv

# 3. Activate virtual environment
source .venv/bin/activate

# 4. Install dependencies
pip install -r requirements.txt
```

### **Configuration Setup**
1. Copy `.env.sample` to `.env`
2. Fill in your API keys and configuration values
3. Ensure required system dependencies are installed (FFmpeg, etc.)

## 📋 Usage Examples

### **Illustration Formatting**
```bash
# Format illustrations for Instagram or fixed dimensions
python image/illustrations_formatter.py -s ./input -d ./output -r 3:4

# Use GUI interface
python image/illustrations_formatter_gui.py
```

### **Screen Recording**
```bash
# Start screen recording
python video/screen_recorder.py

# Or use the batch file
video/screen_recorder.bat
```

### **Notion Newsletter Building**
```bash
# Build newsletter from Notion articles (writes notion/build_newsletter.html, opens browser,
# then prompts for must-read 1/2/3 and copies the ordered title line to the clipboard)
python notion/build_newsletter.py --newsletter 057 --config notion/build_newsletter.json
```

Details: [`notion/build_newsletter.md`](notion/build_newsletter.md).

## ⚙️ Configuration

### **Environment Variables**
Create a `.env` file in the root directory with:
```env
# Notion API
NOTION_API_TOKEN=your_notion_api_token
NOTION_DATABASE_ID=your_database_id

# Audio settings
AUDIO_OUTPUT_PATH=path/to/audio/output
DEFAULT_LANGUAGE=Spanish

# Machine-local folder paths (no hardcoded defaults — must be set per machine)
NOTION_PARAMS_FILE=         # path to the legacy notion-params.txt file (notion/utils.py)
ILLUSTRATIONS_DEST_INSTAGRAM=  # destination folder for Instagram-formatted illustrations (image/illustrations_formatter*.py)
ILLUSTRATIONS_DEST_1920X1080=  # destination folder for 1920×1080 illustrations (image/illustrations_formatter*.py)
CARROUSEL_SOURCE_FOLDER=    # source folder of carrousel PDF files (image/carrousel_pdf_to_jpg.py)
POPPLER_PATH=               # path to Poppler bin dir if not on system PATH (image/carrousel_pdf_to_jpg.py)
COLLAGE_SOURCE_FOLDER=      # source folder for collage images (image/collage_image.py)
TODOIST_SOURCE_FOLDER=      # source folder for Todoist CSV backup files (notion/todoist_migration.py)
```

### **JSON Configuration Files**
Most modules use JSON configuration files for flexible settings:
- `illustrations_formatter_config.json` - Illustration formatting options
- `build_newsletter.json` - Newsletter configuration
- `cleaning_patterns.json` - Text cleaning patterns

## 🔧 Dependencies

### **Core Dependencies**
- **Audio**: `openai-whisper`, `sounddevice`, `scipy`, `pydub`
- **Image**: `Pillow`, `pillow-heif`, `pymediainfo`
- **Video**: `opencv-python`, `mss`, `pyautogui`
- **PDF**: `PyPDF2`, `openpyxl`
- **Data**: `pandas`, `numpy`
- **Notion**: `notion-client`
- **Windows**: `pywin32`

### **System Requirements**
- **FFmpeg**: Required for audio/video processing
- **Windows COM**: Required for Outlook automation
- **Virtual Environment**: Recommended for dependency isolation

## 📊 Output Formats

### **Audio Processing**
- **Transcription**: Text files with timestamps
- **Audio Files**: WAV, MP3, OGG formats
- **Metadata**: JSON files with processing information

### **Image Processing**
- **Formatted Images**: Instagram-compatible aspect ratios
- **Collages**: Combined image layouts
- **Metadata**: Excel files with image information

### **Video Processing**
- **Screen Recordings**: MP4 format with configurable FPS
- **Processed Videos**: Various formats and resolutions
- **Metadata**: JSON files with recording details

### **Data Export**
- **Excel Files**: Structured data with formatting
- **CSV Files**: Comma-separated data exports
- **JSON Files**: Configuration and metadata storage

## 🐛 Troubleshooting

### **Common Issues**

**FFmpeg Not Found:**
```bash
# Windows (Chocolatey)
choco install ffmpeg

# Linux
sudo apt-get install ffmpeg

# macOS
brew install ffmpeg
```

**Virtual Environment Issues:**
```bash
# Check virtual environment status
python system/venv_manager.py

# Recreate virtual environment
rm -rf .venv
python -m venv .venv
```

**Permission Errors:**
- Ensure you have write access to output directories
- Run as administrator if needed (Windows)
- Check file permissions (Unix/Linux)

### **Debug Mode**
Most scripts support debug mode:
```bash
python script.py --debug
```

## 🔒 Security & Best Practices

### **API Key Management**
- Store sensitive data in `.env` files
- Never commit API keys to version control
- Use environment variables for configuration

### **Data Privacy**
- Implement data sanitization for sensitive information
- Use secure file handling practices
- Follow GDPR and privacy regulations

### **Error Handling**
- All modules include comprehensive error handling
- Graceful degradation for non-critical failures
- Detailed logging for debugging

## 📚 Documentation

### **Module Documentation**
Each module includes:
- Detailed docstrings and type hints
- Usage examples and configuration options
- Error handling and troubleshooting guides

### **Code Standards**
- Follows PEP 8 and PEP 20 guidelines
- Comprehensive type hints throughout
- Consistent logging and error handling
- Clear function and variable naming

## 🤝 Contributing

### **Development Setup**
1. Fork the repository
2. Create a feature branch
3. Follow the coding standards in `CLAUDE.md`
4. Include comprehensive documentation
5. Submit a pull request following the conventions in `CLAUDE.md`

### **Code Quality**
- All code must pass linting checks
- Include type hints for all functions
- Add comprehensive docstrings
- Follow the project's error handling patterns

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

- OpenAI Whisper for speech recognition
- Notion API for database integration
- PySimpleGUI for user interface components in `system/keycaster.py` and `system/quickdeck/` *(note: PySimpleGUI is no longer freely available on PyPI — it requires a paid license or registration; these two tools need a separately-obtained PySimpleGUI installation, or can be migrated to the drop-in [FreeSimpleGUI](https://github.com/spyoungtech/FreeSimpleGUI) fork)*
- FFmpeg for audio/video processing

---

*Python version: 3.8+*
*Platform: Windows, Linux, macOS*
