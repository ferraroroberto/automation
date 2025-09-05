# Voice Transcription Tool

A comprehensive voice transcription application with GUI and quick mode support for automated transcription workflows.

## 🚀 Features

- **Full GUI Application**: Complete graphical interface for transcription settings
- **Quick Mode**: Silent background transcription without GUI
- **CUDA GPU Support**: Hardware acceleration for faster processing
- **Multi-language Support**: Spanish and English transcription
- **StreamDeck Integration**: Direct launch capabilities
- **Auto-paste**: Automatic clipboard integration
- **Windows Notifications**: System tray notifications

## 📋 Usage

### GUI Mode (Default)
```powershell
& "$env:VENV_FOLDER\Scripts\python.exe" transcribe_voice_gui.py
```

### Quick Mode (No GUI, Auto-Exit)
```powershell
# Spanish transcription
& "$env:VENV_FOLDER\Scripts\python.exe" transcribe_voice_gui.py --quick-spanish

# English transcription
& "$env:VENV_FOLDER\Scripts\python.exe" transcribe_voice_gui.py --quick-english
```

### Quick Launch Scripts
```powershell
# Regular batch files (automatically find venv path)
.\quick_spanish.bat
.\quick_english.bat

# Stream Deck optimized (visible terminal, clear instructions)
.\quick_spanish_streamdeck.bat
.\quick_english_streamdeck.bat

# PowerShell scripts (automatically find venv path)
.\quick_spanish.ps1
.\quick_english.ps1
```

## 🎯 Quick Mode Behavior

### User Experience Flow
```
User runs: python transcribe_voice_gui.py --quick-spanish
↓
Console: "🎤 Recording in Spanish... (press any key to stop early)"
↓
Console: "⏹️  Recording stopped. Processing..."
↓
Console: "✅ Transcription complete. Text copied to clipboard."
↓
Application exits automatically
```

### Key Features
- ✅ **No GUI Display**: Runs completely in background/terminal
- ✅ **Immediate Recording**: Starts recording as soon as dependencies are ready
- ✅ **Automatic Stop**: Stops after default time (300 seconds) or on keypress
- ✅ **Silent Processing**: No status windows, just console output
- ✅ **Auto-Copy & Exit**: Copies transcription to clipboard and exits automatically
- ✅ **Minimal Output**: Only essential status messages

## 🔧 Configuration

### Language Settings
- **Spanish**: Transcribes and translates to English
- **English**: Transcribes only (no translation)

### Model Settings
- **Base Model**: Used in quick mode for optimal speed/quality balance
- **Other Models**: Tiny, Small, Medium, Large (available in GUI mode)

### Recording Settings
- **Default Duration**: 300 seconds (5 minutes)
- **Early Stop**: Press any key to stop recording early
- **Auto-stop**: Recording stops automatically when time limit reached

## 🖥️ System Requirements

### Hardware
- **Microphone**: Any Windows-compatible audio input device
- **GPU**: NVIDIA GeForce GTX 1070+ (recommended for CUDA acceleration)
- **RAM**: 8GB+ recommended

### Software
- **Python**: 3.8+
- **Virtual Environment**: Automatically detected at `E:\automation\automation\.venv`
- **FFmpeg**: Required for audio processing
- **CUDA**: Optional but recommended for GPU acceleration

### Dependencies
All required packages are listed in the main `requirements.txt`:
- `openai-whisper` - Speech-to-text transcription
- `sounddevice` - Audio recording
- `scipy` - Audio processing
- `pynput` - Keyboard monitoring
- `pyperclip` - Clipboard operations
- `torch` - Machine learning framework

## 🎮 StreamDeck Integration

### Quick Mode for Stream Deck (Interactive Terminal)
For Stream Deck buttons that need visible terminal interaction:

```powershell
# Stream Deck optimized launchers (recommended)
.\quick_spanish_streamdeck.bat
.\quick_english_streamdeck.bat
```

**Stream Deck Features:**
- ✅ **Visible Terminal**: Window appears on top and focused
- ✅ **Clear Instructions**: Shows what to do and how to stop
- ✅ **Interactive**: Press any key to stop recording early
- ✅ **Auto-focus**: Window automatically comes to foreground

### Direct Launch Commands (Legacy)
```powershell
# Spanish with focus priority
& "$env:VENV_FOLDER\Scripts\python.exe" transcribe_voice_gui.py --direct_record_spanish --focus_priority 3

# English with focus priority
& "$env:VENV_FOLDER\Scripts\python.exe" transcribe_voice_gui.py --direct_record_english --focus_priority 3
```

## 🛠️ Development

### File Structure
```
transcribe_voice/
├── transcribe_voice_gui.py          # Main GUI application with quick mode
├── transcribe_voice_core.py         # Core transcription functionality
├── quick_spanish.bat               # Quick Spanish launcher (batch)
├── quick_english.bat               # Quick English launcher (batch)
├── quick_spanish_streamdeck.bat    # Stream Deck Spanish launcher
├── quick_english_streamdeck.bat    # Stream Deck English launcher
├── quick_spanish.ps1               # Quick Spanish launcher (PowerShell)
├── quick_english.ps1               # Quick English launcher (PowerShell)
├── CUDA.md                         # CUDA setup guide
├── README_CUDA.md                  # CUDA quick reference
└── README.md                       # This file
```

### Adding New Languages
1. Update language selection in GUI
2. Add quick mode argument in `main()`
3. Create corresponding launcher scripts
4. Update documentation

## 🔧 Troubleshooting

### Common Issues

#### "Module not found" errors
```powershell
# Activate virtual environment and install dependencies
& "$env:VENV_FOLDER\Scripts\python.exe" -m pip install -r requirements.txt
```

#### No audio devices found
- Check microphone connections
- Verify microphone permissions in Windows settings
- Test with different audio devices

#### CUDA not working
- Follow `CUDA.md` for proper installation
- Verify GPU drivers are up to date
- Check CUDA compatibility with your GPU

#### Recording too short
- Check audio levels during recording
- Verify microphone sensitivity
- Test with different audio devices

### Error Messages
- **"❌ ERROR: Required dependencies not available"**: Install missing packages
- **"❌ ERROR: No input devices found!"**: Check microphone setup
- **"❌ ERROR: No suitable microphone found!"**: Verify preferred microphone configuration

## 📚 API Reference

### Command Line Arguments
- `--quick-spanish`: Launch quick mode for Spanish
- `--quick-english`: Launch quick mode for English
- `--launch_from_elgato`: StreamDeck compatibility mode
- `--direct_record_spanish`: Direct Spanish recording (legacy)
- `--direct_record_english`: Direct English recording (legacy)

### Configuration Classes
- `TranscriptionConfig`: Main configuration settings
- `AudioRecorder`: Audio recording functionality
- `Transcriber`: Speech-to-text processing

## 🔄 Updates and Maintenance

### Virtual Environment
Always use the configured virtual environment:
```powershell
# Check VENV_FOLDER in .env file
$VenvPath = (Get-Content .env | Select-String "VENV_FOLDER" | ForEach-Object { $_.Line -split "=" | Select-Object -Last 1 }).Trim()

# Use virtual environment Python
& "$VenvPath\Scripts\python.exe" transcribe_voice_gui.py --quick-spanish
```

### Dependency Updates
```powershell
# Update all packages
& "$env:VENV_FOLDER\Scripts\python.exe" -m pip install --upgrade -r requirements.txt

# Update specific package
& "$env:VENV_FOLDER\Scripts\python.exe" -m pip install --upgrade openai-whisper
```

## 📝 License and Credits

This tool uses:
- **OpenAI Whisper**: For speech-to-text transcription
- **PyTorch**: Machine learning framework
- **SoundDevice**: Audio recording
- **CUDA**: GPU acceleration (optional)

## 🆘 Support

For issues and questions:
1. Check this README first
2. Review `CUDA.md` for GPU-related issues
3. Test with the provided batch/PowerShell scripts
4. Verify virtual environment configuration

---

**Quick Start Commands:**
```powershell
# Spanish quick transcription
.\quick_spanish.ps1
# or for Stream Deck (visible terminal)
.\quick_spanish_streamdeck.bat

# English quick transcription
.\quick_english.ps1
# or for Stream Deck (visible terminal)
.\quick_english_streamdeck.bat

# Full GUI mode
& "$env:VENV_FOLDER\Scripts\python.exe" transcribe_voice_gui.py
```
