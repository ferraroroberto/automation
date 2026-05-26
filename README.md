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
├── 📁 launcher/        # Flask remote launcher (phone → host via Tailscale)
├── 📁 linkedin/        # LinkedIn automation, IP checking, and profile data extraction
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
- **`audio_normalize.py`** - Audio normalization

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
- **`build_newsletter.py`** - Newsletter builder from Notion articles — see [`notion/build_newsletter.md`](notion/build_newsletter.md)
- **`notion_databases_dump.py`** - Export Notion databases to Excel
- **`notion_databases_clean.py`** - Clean and normalize Notion data
- **`notion_databases_add_editorial.py`** - Add editorial calendar rows (one day per date) to a Notion database — see [`notion/notion_databases_add_editorial.md`](notion/notion_databases_add_editorial.md)
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
- **`wifi/`** - Wi‑Fi connection scripts, saved‑password export, BAT generator GUI, and network device scanner — see [`system/wifi/README.md`](system/wifi/README.md)
- **`apple/`** - Apple-related conversion and utilities

**File & Document Operations**
- **`unzip_with_password.py`** - Password-protected archive extraction
- **`rename_files.py`** - Batch file renaming utilities
- **`copy_git_project.py`** - Git project copying and setup
- **`base64_encode_decode.py`** - Base64 encode/decode utility
- **`word_to_markdown.py`** - Word to Markdown conversion
- **`foldersearcher/`**, **`treesize/`** - Folder search and size utilities

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

### 📱 Remote Launcher (`launcher/`)

A Flask web UI (password optional) that lets you start projects on the host machine from your phone over Tailscale.

**`launcher.py`** — discovers every `*remote*.bat` in the parent directory, presents them as buttons, and spawns the selected one in a new CMD window on the host. A **Generate BAT files** page (`/generate`) keeps `*-remote.bat` files in sync with `.code-workspace` files — add a workspace, hit Generate, done.

#### Running the launcher

**Headless (terminal only):**
```powershell
& .\.venv\Scripts\python.exe launcher\launcher.py
```

**System-tray GUI** (server starts automatically; double-click the icon to open the status + log window):
```powershell
& .\.venv\Scripts\python.exe launcher\tray.py
```

Required `.env` keys:

```env
LAUNCHER_PASSWORD=your_login_password
LAUNCHER_SECRET_KEY=a_long_random_string

# Optional overrides
LAUNCHER_PROJECTS_DIR=C:\path\to\bat\files   # default: parent of this repo
LAUNCHER_PORT=5050
LAUNCHER_HOST=0.0.0.0
```

#### Enabling HTTPS via Tailscale (recommended)

Tailscale issues free, browser-trusted TLS certificates for every device on your tailnet. Follow these steps once; the whole process takes about five minutes.

---

**Step 1 — Enable MagicDNS in the Tailscale admin console**

1. Open <https://login.tailscale.com/admin/dns> in a browser.
2. Under the **Nameservers** section, look for the **MagicDNS** toggle and switch it on.
   - MagicDNS gives every device on your tailnet a stable `*.ts.net` hostname instead of a bare IP address.

---

**Step 2 — Enable HTTPS Certificates**

On the same DNS settings page, scroll down to the **HTTPS Certificates** section and click **Enable HTTPS**.

> If you don't see this section, your account may need to be on a paid plan — check the Tailscale pricing page.

---

**Step 3 — Find your device's MagicDNS hostname**

Run this in a PowerShell or CMD window on the host machine:

```powershell
tailscale status
```

Look for the line that starts with your machine's name. The full hostname is in the last column and ends in `.ts.net`, e.g.:

```
100.x.y.z   your-pc   your-account@  windows  -
```

The MagicDNS hostname will be `your-pc.tail1234.ts.net` (the exact middle segment is your tailnet name, visible in the Tailscale admin under **Settings → General**).

Alternatively, run:

```powershell
tailscale status --json | python -c "import sys,json; s=json.load(sys.stdin); print(s['Self']['DNSName'].rstrip('.'))"
```

This prints the full hostname directly, e.g. `your-pc.tail1234.ts.net`.

---

**Step 4 — Create the `certificates/` folder and request the certificate**

Create the folder inside the repo root (it is already in `.gitignore`):

```powershell
mkdir certificates
```

Then request the certificate — Tailscale writes both files into the current directory, so run from that folder:

```powershell
cd certificates
tailscale cert your-pc.tail1234.ts.net
```

> **Important:** Steps 1 and 2 (MagicDNS + HTTPS Certificates) must be completed in the admin console *before* running this command. If you run `tailscale cert` first, Tailscale issues a self-signed certificate that browsers will reject with an "untrusted" or "not secure" warning. If that happened, delete both files in `certificates/` and re-run `tailscale cert` now that HTTPS is enabled.

You should see output like:

```
Wrote public cert to your-pc.tail1234.ts.net.crt
Wrote private key to your-pc.tail1234.ts.net.key
```

Two files are now in `certificates/`:

| File | Contents |
|------|----------|
| `your-pc.tail1234.ts.net.crt` | Certificate chain (public, safe to share) |
| `your-pc.tail1234.ts.net.key` | Private key — **keep this secret** |

> If `tailscale cert` is not found, Tailscale may not be on your PATH. Try the full path, usually `"C:\Program Files\Tailscale\tailscale.exe" cert ...`

---

**Step 5 — Add the paths to `.env`**

Open the root `.env` file and add (use your actual hostname):

```env
LAUNCHER_SSL_CERT=certificates\your-pc.tail1234.ts.net.crt
LAUNCHER_SSL_KEY=certificates\your-pc.tail1234.ts.net.key
```

Paths are relative to where you run the launcher from (the repo root). Absolute paths work too.

> Both variables must be set together. Setting only one causes the launcher to exit immediately with an error message.

---

**Step 6 — Start the launcher and verify**

```powershell
& .\.venv\Scripts\python.exe launcher\tray.py
```

The tray icon turns green and the log shows:

```
ℹ️ Launcher serving on https://0.0.0.0:5050
```

Open `https://your-pc.tail1234.ts.net:5050` from your phone (connected to the same Tailscale network). The browser should show a padlock with no warnings. If it warns about an untrusted certificate, double-check that HTTPS was enabled in Step 2 before you ran `tailscale cert`.

---

**Certificate renewal**

Tailscale certificates expire after **~90 days**. To renew, `cd` into the `certificates/` folder and re-run `tailscale cert your-pc.tail1234.ts.net` — it overwrites the files in place. Restart the launcher to pick up the new certificate.

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
# Build newsletter from Notion articles (writes notion/build_newsletter.html, opens browser,
# then prompts for must-read 1/2/3 and copies the ordered title line to the clipboard)
python notion/build_newsletter.py --newsletter 057 --config notion/build_newsletter.json
```

Details: [`notion/build_newsletter.md`](notion/build_newsletter.md).

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
- PySimpleGUI for user interface components
- FFmpeg for audio/video processing

---

*Last updated: February 2026*
*Python version: 3.8+*
*Platform: Windows, Linux, macOS*
