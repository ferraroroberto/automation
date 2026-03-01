# 🚀 Automation Tools Collection

A comprehensive collection of Python automation tools for audio processing, image manipulation, video recording, email management, Notion integration, and system utilities.

## 🎯 Overview

This repository contains a diverse set of automation tools designed to streamline common digital tasks. Each module is built with a focus on reliability, user experience, and cross-platform compatibility. The tools follow consistent coding standards with proper error handling, logging, and configuration management.

## 🏗️ Project Structure

```
automation/
├── 📁 audio/           # Audio recording, transcription, and conversion tools
├── 📁 email/           # Outlook automation and email management
├── 📁 excel/           # Excel automation (e.g. Stripe accounting)
├── 📁 google/          # Gmail, Drive, and Google Photos automation
├── 📁 html/            # HTML utilities (e.g. countdown timer)
├── 📁 image/           # Image processing, formatting, and Instagram tools
├── 📁 linkedin/        # LinkedIn automation, IP checking, and profile data extraction
├── 📁 notion/          # Notion API integration and database management
├── 📁 smart_life/      # Smart Life / IoT device automation
├── 📁 system/          # System utilities and virtual environment management
├── 📁 text/            # Text processing and PDF conversion tools
└── 📁 video/           # Screen recording and video processing tools
```

Configuration uses a root `.env` file (see [Configuration](#️-configuration)); copy from `.env.sample` if present.

### 📚 Reference Documents

- **`AGENTS.md`** - Master Onboarding & Context Map (Read this first!)
- **`AGENTS_PYTHON.md`** - Python Coding Standards & Patterns
- **`AGENTS_POWERSHELL.md`** - PowerShell & Shell Standards
- **`AGENTS_STRUCTURE.md`** - Project Structure & Refactoring Rules
- **`AGENTS_CLI.md`** - Guide for building CLI tools
- **`AGENTS_PR.md`** - Template for creating pull requests

> Guidelines adapted from [HumanLayer's "Writing a good CLAUDE.md"](https://www.humanlayer.dev/blog/writing-a-good-claude-md).

## 🔧 Core Modules

### 🎵 Audio Processing (`audio/`)

**Voice Transcription & Recording**
- **`transcribe_voice/`** - Core recording and OpenAI Whisper transcription (`transcribe_voice_core.py`), GUI launcher, and batch helpers
- **`audio_extractor_core.py`** / **`audio_extractor_gui.py`** - Audio extraction and processing
- **`convert_ogg_mp3.py`** - Audio format conversion (OGG to MP3)
- **`transcript_collate.py`** - Collate and manage transcripts
- **`audio_normalize.py`** - Audio normalization

**Features:**
- Real-time audio recording with configurable duration
- GPU acceleration support for faster transcription
- Multi-language support with translation capabilities
- Automatic microphone detection and selection
- Temporary file management and cleanup

### 🖼️ Image Processing (`image/`)

**Instagram & Social Media Tools**
- **`instagram_formatter.py`** - Convert images to Instagram-compatible formats
- **`instagram_formatter_gui.py`** - GUI for Instagram image formatting
- **`photos_archive.py`** - Photo organization and metadata management
- **`image_resizer.py`** - Batch image resizing and optimization

**PDF & Document Processing**
- **`carrousel_pdf_to_jpg.py`** - Convert PDF pages to JPG images
- **`png_to_jpg_converter.py`** - PNG to JPG format conversion

**Utilities**
- **`screenshot.py`** - Automated screenshot capture
- **`collage_image.py`** - Create image collages from multiple files
- **`gif_unpacker.py`** - Extract frames from GIFs
- **`illustrations_formatter.py`** - Illustration formatting and batch processing
- **`transparency_variants.py`** - Generate transparency variants
- **`pdf_to_jpg/`** - PDF to JPG conversion (see `README_PDF_Converter.md`)

### 📹 Video Processing (`video/`)

**Screen Recording & Capture**
- **`screen_recorder.py`** - Multi-monitor screen recording with mouse overlay
- **`video_trim.py`** - Video trimming and editing utilities
- **`video_concatenator.py`** - Combine multiple video files

**Download & Processing**
- **`video_download.py`** - Unified video downloader (YouTube, HLS/M3U8, direct URL) with GUI
- **`video_reencoder.py`** - Video format conversion and re-encoding

### 📧 Email Automation (`email/`)

**Outlook Integration**
- **`email-automation-save.py`** - Save selected emails with attachments
- **`email-automation-archive.py`** - Email archiving and organization
- **`email-automation-classify.py`** - Email classification and sorting
- **`utils.py`** - Common email processing utilities

**Features:**
- Outlook COM automation for Windows
- Attachment extraction and organization
- Excel-based email tracking and metadata
- Gmail archive integration

### 📚 Notion Integration (`notion/`)

**Database Management**
- **`build_newsletter.py`** - Newsletter builder from Notion articles
- **`notion_databases_dump.py`** - Export Notion databases to Excel
- **`notion_databases_clean.py`** - Clean and normalize Notion data
- **`notion_databases_add_editorial.py`** - Add editorial metadata to Notion databases
- **`notion_excel_sync.py`** - Excel-Notion synchronization
- **`articles_sync/`** - Notion articles sync and incremental sync

**Content Processing**
- **`normalize_names.py`** - Name normalization and standardization
- **`todoist_migration.py`** - Todoist to Notion migration tool
- **`journal_automation.py`** - Automated journal entry creation

### 💻 System Utilities (`system/`)

**Environment & Input**
- **`venv_manager.py`** - Virtual environment management and monitoring
- **`mouse_mover.py`** - Automated mouse movement and clicking
- **`keycaster.py`** - Keyboard input automation and monitoring
- **`quickdeck/`** - Quick Deck / Stream Deck integration
- **`textexpander/`** - Text expander and prompt templates
- **`wifi/`** - Wi‑Fi password retrieval and connection scripts
- **`apple/`** - Apple-related conversion and utilities

**File & Document Operations**
- **`unzip_with_password.py`** - Password-protected archive extraction
- **`rename_files.py`** - Batch file renaming utilities
- **`copy_git_project.py`** - Git project copying and setup
- **`base64_encode_decode.py`** - Base64 encode/decode utility
- **`word_to_markdown.py`** - Word to Markdown conversion
- **`foldersearcher/`**, **`treesize/`** - Folder search and size utilities
- **`grocery/`** - Grocery list / app (Streamlit)

### 📝 Text Processing (`text/`)

**Document Conversion**
- **`convert_pdf_to_txt.py`** - PDF to text extraction
- **`clean_sensitive_data.py`** - Sensitive data removal and sanitization

### 🔗 LinkedIn Tools (`linkedin/`)

**IP & Profile Automation**
- **`check_ip/`** - IP address checking and management tools
- **`open_profiles/`** - LinkedIn profile opening automation
- **`profiles_data_extractor/`** - DevTools-based profile data extraction, Streamlit dashboard, Excel formatting, and data entry tools

### ☁️ Google Integration (`google/`)

**Gmail, Drive & Photos**
- **`gmail_drive_automation.py`** - Gmail and Google Drive automation
- **`weekly_photo_automation.py`** - Weekly photo automation (e.g. Google Photos)
- **`setup_helper_script.py`** - Setup and configuration helpers

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

### **Voice Transcription**
```bash
# Start voice transcription with GUI
python audio/transcribe_voice_gui.py

# Command-line transcription
python audio/transcribe_voice_core.py --record-seconds 300 --language Spanish
```

### **Instagram Image Formatting**
```bash
# Format images for Instagram
python image/instagram_formatter.py --source-folder ./input --aspect-ratio 3:4

# Use GUI interface
python image/instagram_formatter_gui.py
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
# Build newsletter from Notion articles
python notion/build_newsletter.py --newsletter "057" --config "build_newsletter.json"
```

### **Email Automation**
```bash
# Save selected Outlook email
python email/email-automation-save.py

# Archive emails
python email/email-automation-archive.py
```

## ⚙️ Configuration

### **Environment Variables**
Create a `.env` file in the root directory with:
```env
# Notion API
NOTION_API_KEY=your_notion_api_key
NOTION_DATABASE_ID=your_database_id

# Email settings
EMAIL_FOLDER_PATH=path/to/email/folder
EXCEL_TRACKING_PATH=path/to/tracking.xlsx

# Audio settings
AUDIO_OUTPUT_PATH=path/to/audio/output
DEFAULT_LANGUAGE=Spanish
```

### **JSON Configuration Files**
Most modules use JSON configuration files for flexible settings:
- `instagram_formatter_config.json` - Instagram formatting options
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
3. Follow the coding standards in `AGENTS_PYTHON.md` and related `AGENTS_*.md` docs
4. Include comprehensive documentation
5. Submit a pull request using the template in `AGENTS_PR.md`

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
- PySimpleGUI for user interface components
- FFmpeg for audio/video processing

---

*Last updated: February 2026*
*Python version: 3.8+*
*Platform: Windows, Linux, macOS*
