# 🚀 Automation Tools Collection

A comprehensive collection of Python automation tools for audio processing, image manipulation, video recording, email management, Notion integration, and system utilities.

## 🎯 Overview

This repository contains a diverse set of automation tools designed to streamline common digital tasks. Each folder is an independent set of scripts rather than part of one application — most target Windows, a few are cross-platform, and they vary in maturity. Shared conventions (secrets, verification, branch pipeline) are listed under [Conventions](#-conventions).

## 🏗️ Project Structure

```
automation/
├── 📁 docs/            # Durable reference docs (see architecture.mmd below)
├── 📁 scripts/         # Repo-level scripts (verification gate)
├── 📁 tests/           # Every test, under tests/<domain>/
└── 📁 tools/           # One folder per tool domain
    ├── 📁 audio/       # Audio recording, transcription, and conversion tools
    ├── 📁 excel/       # Excel automation (e.g. Stripe accounting)
    ├── 📁 google/      # Gmail, Drive, and Google Photos automation
    ├── 📁 html/        # HTML utilities (e.g. countdown timer)
    ├── 📁 image/       # Image processing, formatting, and Instagram tools
    ├── 📁 linkedin/    # LinkedIn reverse-image search, profile opening, and profile data extraction
    ├── 📁 notion/      # Notion API integration and database management
    ├── 📁 parking/     # Parking reservation poller (15-min heartbeat, books a free slot)
    ├── 📁 smart_life/  # Smart Life / IoT device automation
    ├── 📁 system/      # System utilities and virtual environment management
    ├── 📁 text/        # Text processing and PDF conversion tools
    └── 📁 video/       # Screen recording and video processing tools
```

[`docs/layout.md`](docs/layout.md) says what each folder and subpackage holds and where new code and tests go; the structure ratchet in the gate enforces it.

[`docs/architecture.mmd`](docs/architecture.mmd) is the hand-authored Mermaid diagram of how those folders relate — the domain folders, the shared per-domain helpers (`tools/google/_auth.py`, `tools/notion/utils.py`), and the external services each domain talks to. `CLAUDE.md` requires it be updated in the same PR as any structural change.

Configuration uses a root `.env` file (see [Configuration](#️-configuration)); copy from `.env.sample` if present.

### 📚 Agent Instructions

- **`CLAUDE.md`** — canonical instruction set for any AI coding agent working in this repo.
- **`AGENTS.md`** — one-line pointer to `CLAUDE.md` for non-Claude agents (Cursor, Codex, etc.).

The master template lives in [`project-scaffolding/docs/agents/`](../project-scaffolding/docs/agents/) and is propagated to every sibling repo.

## 🔧 Core Modules

### 🎵 Audio Processing (`tools/audio/`)

- **`audio_extractor_core.py`** / **`audio_extractor_gui.py`** - Audio extraction and processing
- **`convert_ogg_mp3.py`** - Audio format conversion (OGG to MP3)
- **`transcript_collate.py`** - Collate and manage transcripts
- **`audio_normalize.py`** - Audio normalization (EBU R128 loudness)
- **`audio_trim.py`** - Interactive audio trimming

Full per-script reference: [`tools/audio/README.md`](tools/audio/README.md).

### 🖼️ Image Processing (`tools/image/`)

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

Reference for the previously-undocumented tools above: [`tools/image/README.md`](tools/image/README.md).

### 📹 Video Processing (`tools/video/`)

**Screen Recording & Capture**
- **`screen_recorder.py`** - Multi-monitor screen recording with mouse overlay — see [`tools/video/screen_recorder.md`](tools/video/screen_recorder.md)
- **`video_trim.py`** - Video trimming and editing utilities
- **`video_concatenator.py`** - Combine multiple video files

**Download & Processing**
- **`video_download.py`** - Unified video downloader (YouTube, HLS/M3U8, direct URL) with GUI
- **`video_reencoder.py`** - Video format conversion and re-encoding

**GPU Maintenance**
- **`gpu_recovery.py`** - Diagnose & recover an NVIDIA GPU stuck "disabled" (Device Manager Code 22); restores NVENC acceleration for the tools above (see `gpu_recovery.md`)

### 🅿️ Parking Reservation (`tools/parking/`)

A deterministic poller (no LLM) that checks every 15 minutes for a free parking slot on the configured weekdays, books one day per call, and sends a Telegram message. Runs as the app-launcher Jobs entry `parking-poll`; dry-run by default. Usage, config and the monthly login: [`tools/parking/README.md`](tools/parking/README.md).

### 📚 Notion Integration (`tools/notion/`)

**Database Management**
- **`build_newsletter.py`** - Newsletter builder from Notion articles — see [`tools/notion/build_newsletter.md`](tools/notion/build_newsletter.md)
- **`notion_databases_dump.py`** - Export Notion databases to Excel
- **`notion_databases_query.py`** - List every database the integration can see and export name/id/url to Excel — see [`tools/notion/notion_databases_query.md`](tools/notion/notion_databases_query.md)
- **`notion_databases_clean.py`** - Apply per-column extract/keep/rename/reorder rules to dumped databases — see [`tools/notion/notion_databases_clean.md`](tools/notion/notion_databases_clean.md)
- **`notion_databases_add_editorial.py`** - Add editorial calendar rows (one day per date) to a Notion database — see [`tools/notion/notion_databases_add_editorial.md`](tools/notion/notion_databases_add_editorial.md)
- **`notion_databases_editorial.py`** - Local Excel-to-Excel refresh of the editorial-calendar workbook (no Notion API calls) — see [`tools/notion/notion_databases_editorial.md`](tools/notion/notion_databases_editorial.md)
- **`notion_excel_sync.py`** - Push rows from an Excel workbook into a Notion database as new pages — see [`tools/notion/notion_excel_sync.md`](tools/notion/notion_excel_sync.md)
- **`sample_illustrations.py`** - Sample illustrations from a Notion database — see [`tools/notion/sample_illustrations.md`](tools/notion/sample_illustrations.md)
- **`articles_sync/`** - Notion articles sync and incremental sync — see [`tools/notion/articles_sync/notion_articles_sync_readme.md`](tools/notion/articles_sync/notion_articles_sync_readme.md)

**Content Processing**
- **`normalize_names.py`** - Name normalization and standardization — see [`tools/notion/normalize_names.md`](tools/notion/normalize_names.md)
- **`normalize_url.py`** - Strip tracking params from URLs in a Notion database — see [`tools/notion/normalize_url.md`](tools/notion/normalize_url.md)
- **`todoist_migration.py`** - Transform a folder of Todoist CSV exports into a consolidated Excel workbook — see [`tools/notion/todoist_migration.md`](tools/notion/todoist_migration.md)
- **`journal_automation.py`** - Automated journal entry creation — see [`tools/notion/journal_automation.md`](tools/notion/journal_automation.md)

### 💻 System Utilities (`tools/system/`)

**Environment & Input**
- **`venv_manager.py`** - Virtual environment management and monitoring
- **`mouse_mover.py`** - Automated mouse movement and clicking
- **`keycaster.py`** - Keyboard input automation and monitoring
- **`quickdeck/`** - Quick Deck / Stream Deck integration
- **`textexpander/`** - Text expander and prompt templates
- **`wifi/`** - Wi‑Fi connection scripts, saved‑password export, BAT generator GUI, and network device scanner — see [`tools/system/wifi/README.md`](tools/system/wifi/README.md)
- **`apple/`** - vCard tools: `convert_to_apple.py` (GUI converter to Apple-compatible vCard 3.0) and `unify_vcards.py` (interactive duplicate-contact merging) — see [`tools/system/apple/convert_to_apple.md`](tools/system/apple/convert_to_apple.md)

**File & Document Operations**
- **`unzip_with_password.py`** - Password-protected archive extraction
- **`rename_files.py`** - Batch file renaming utilities
- **`copy_git_project.py`** - Git project copying and setup
- **`base64_encode_decode.py`** - Base64 encode/decode utility
- **`word_to_markdown.py`** - Word to Markdown conversion
- **`list_files_to_xls.py`** - List a folder's files to an Excel workbook
- **`background.py`** - Set desktop/taskbar color theme (black or light grey)
- **`outlook_web_mail.py`** / **`outlook_web.py`** - Playwright tool and reusable library that draft (and, only when asked, send) Outlook-on-the-web mail with any files attached, incl. numbered parts of a split file — see [`tools/system/outlook_web_mail.md`](tools/system/outlook_web_mail.md)
- **`markdown_preview.py`** - Tray-resident GitHub-style Markdown previewer (Edge WebView2) with light/dark toggle and live reload — see [`tools/system/markdown_preview.md`](tools/system/markdown_preview.md)
- **`foldersearcher/`** - Retired (folder-explorer's folder search replaces it): tray-resident multi-root folder-name search with email-branch pruning and path-depth display, plus a headless `scan` CLI — see [`tools/system/foldersearcher/foldersearcher.md`](tools/system/foldersearcher/foldersearcher.md)
- **`treesize/`** - Folder size utilities

Reference for the previously-undocumented `tools/system/` scripts (background, keycaster, unzip_with_password, rename_files, copy_git_project, word_to_markdown, list_files_to_xls): [`tools/system/README.md`](tools/system/README.md).

### 📝 Text Processing (`tools/text/`)

**Document Conversion**
- **`convert_pdf_to_txt.py`** - PDF to text extraction
- **`clean_sensitive_data.py`** - Sensitive data removal and sanitization

Full per-script reference: [`tools/text/README.md`](tools/text/README.md).

### 🔗 LinkedIn Tools (`tools/linkedin/`)

**Reverse Image Search & Profile Automation**
- **`check_ip/`** - Reverse-image search (Google Lens via SerpAPI) that uploads images to Imgur and counts where they appear across LinkedIn and other social platforms — see [`tools/linkedin/check_ip/README.md`](tools/linkedin/check_ip/README.md)
- **`open_profiles/`** - Opens a filtered set of LinkedIn profile/company activity pages from an Excel export — see [`tools/linkedin/open_profiles/README.md`](tools/linkedin/open_profiles/README.md)
- **`profiles_data_extractor/`** - DevTools-based profile data extraction, Streamlit dashboard, Excel formatting, and data entry tools

### ☁️ Google Integration (`tools/google/`)

**Gmail, Drive & Photos**
- **`gmail_drive_automation.py`** - Gmail and Google Drive automation — see [`tools/google/gmail_drive_automation.md`](tools/google/gmail_drive_automation.md)
- **`photos_weekly_album.py`** - Builds the weekly Google Photos album, shares it by link and drafts the email to a configured list through the web UIs (Playwright, real Chrome; sends only with `--send`); `photos_weekly_album_job.py` runs it as the Saturday-morning app-launcher job and sends a Telegram message with the outcome — see [`tools/google/photos_weekly_album.md`](tools/google/photos_weekly_album.md)
- **`diag_gmail_drive.py`** - Offline diagnostic for the Gmail/Drive automation — see [`tools/google/README.md`](tools/google/README.md)

### 📊 Excel Automation (`tools/excel/`)

- **`accounting_stripe/`** - Stripe accounting and Excel workflows

### 🌐 HTML Utilities (`tools/html/`)

- **`countdown_timer.html`** - Countdown timer and similar utilities

### 🏠 Smart Life / IoT (`tools/smart_life/`)

- Device automation (e.g. **`despacho_switch.bat`**) and configuration via `devices.sample.json`

## 🚀 Quick Start

### **Prerequisites**
- Python 3.9+ installed — the floor is set by `requirements.txt`, whose pinned `en-core-web-sm` 3.8.0 wheel requires spaCy ≥ 3.8 (`Requires-Python >=3.9`). Developed and verified on 3.14.
- FFmpeg for audio/video processing
- Git for repository management

### **Environment Setup**

**Windows PowerShell:**
```powershell
# 1. Clone the repository
git clone https://github.com/ferraroroberto/automation.git
cd automation

# 2. Create the virtual environment once
python -m venv .venv

# 3. Install dependencies — invoke the venv's python by path, never activate
& .\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

**Unix/Linux/macOS:**
```bash
# 1. Clone the repository
git clone https://github.com/ferraroroberto/automation.git
cd automation

# 2. Create the virtual environment once
python3 -m venv .venv

# 3. Install dependencies — invoke the venv's python by path, never activate
./.venv/bin/python -m pip install -r requirements.txt
```

### **Configuration Setup**
1. Copy `.env.sample` to `.env`
2. Fill in your API keys and configuration values
3. Ensure required system dependencies are installed (FFmpeg, etc.)

## 📋 Usage Examples

### **Illustration Formatting**
```bash
# Format illustrations for Instagram or fixed dimensions
.\.venv\Scripts\python.exe tools/image/illustrations_formatter.py -s ./input -d ./output -r 3:4

# Use GUI interface
.\.venv\Scripts\python.exe tools/image/illustrations_formatter_gui.py
```

### **Screen Recording**
```bash
# Start screen recording
.\.venv\Scripts\python.exe tools/video/screen_recorder.py

# Or use the batch file
tools/video/screen_recorder.bat
```

### **Notion Newsletter Building**
```bash
# Build newsletter from Notion articles (writes tools/notion/build_newsletter.html, opens browser,
# then prompts for must-read 1/2/3 and copies the ordered title line to the clipboard)
.\.venv\Scripts\python.exe tools/notion/build_newsletter.py --newsletter 057 --config tools/notion/build_newsletter.json
```

Details: [`tools/notion/build_newsletter.md`](tools/notion/build_newsletter.md).

## ⚙️ Configuration

### **Environment Variables**
[`.env.sample`](.env.sample) is the single authoritative list of keys — copy it to `.env` and fill in the values. Covers the Notion, Imgur, SerpAPI and Tuya API credentials and the machine-local folder paths (no hardcoded defaults, must be set per machine).

### **JSON Configuration Files**
Most modules use JSON configuration files for flexible settings. A config that holds machine-local paths is gitignored and tracked only as a `<name>.json.sample` template — copy the `.sample` next to it (dropping the `.sample` suffix) and edit the paths for your machine:
- `illustrations_formatter_config.json` - Illustration formatting options (local; template `illustrations_formatter_config.json.sample`)
- `build_newsletter.json` - Newsletter configuration
- `cleaning_patterns.json` - Text cleaning patterns

## 🔧 Dependencies

### **Core Dependencies**
- **Audio**: `openai-whisper`, `scipy`, `imageio-ffmpeg`
- **Image**: `Pillow`, `pillow-heif`, `pymediainfo`
- **Video**: `opencv-python`, `mss`, `pyautogui`
- **PDF**: `pypdf`, `openpyxl`
- **Data**: `pandas`, `numpy`
- **Notion**: `notion-client`
- **Windows**: `pywin32`

### **System Requirements**
- **FFmpeg**: Required for audio/video processing
- **Virtual Environment**: Recommended for dependency isolation

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
& .\.venv\Scripts\python.exe tools/system/venv_manager.py

# Recreate the virtual environment
rm -rf .venv
python -m venv .venv
& .\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

**Permission Errors:**
- Ensure you have write access to output directories
- Run as administrator if needed (Windows)
- Check file permissions (Unix/Linux)

### **Debug Mode**
12 of the ~156 Python scripts accept a `--debug` flag (`grep -rl -- --debug --include=*.py .` to list them); most don't — check a script's own `--help` before assuming it.

## 🔒 Secrets

- API keys, tokens, and machine-local paths live in the root `.env` (git-ignored); `.env.sample` documents the keys.
- `client_secret.json` and `token.json` under `tools/google/` are credentials — never commit them.
- Git author email is pinned to the GitHub noreply address; see [`docs/git-author-email.md`](docs/git-author-email.md).

## 🧪 Conventions

This is a grab-bag of independent scripts, not a single application — there is no linter config and no CI workflow. What the repo does hold itself to:

- **Verification before shipping:** `powershell -File scripts\verify-before-ship.ps1` — byte-compiles every module, runs the structure, instruction, doc refs, dead code, static, docs, comment and conventions ratchets, then the unit-test suites under `tests/`. See [`CLAUDE.md`](CLAUDE.md) § Verification for the current list; the script is the source of truth. It must exit 0.
- **The repo `.venv`, by path:** `& .\.venv\Scripts\python.exe ...` — a bare `python`/`py` is not reliably on PATH on this machine.
- **Branch-based pipeline** (no forks): one issue → one `<type>/<issue-N>-<slug>` branch → one PR → squash-merge. Never commit to `main` directly. Full rules in [`CLAUDE.md`](CLAUDE.md).
- **New or moved domain folders and shared helpers** update [`docs/architecture.mmd`](docs/architecture.mmd) in the same PR.

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

- OpenAI Whisper for speech recognition
- Notion API for database integration
- PySimpleGUI for user interface components in `tools/system/keycaster.py` and `tools/system/quickdeck/` *(note: PySimpleGUI is no longer freely available on PyPI — it requires a paid license or registration; these two tools need a separately-obtained PySimpleGUI installation, or can be migrated to the drop-in [FreeSimpleGUI](https://github.com/spyoungtech/FreeSimpleGUI) fork)*
- FFmpeg for audio/video processing

---

*Python version: 3.9+ (developed on 3.14)*
*Platform: Windows, Linux, macOS*
