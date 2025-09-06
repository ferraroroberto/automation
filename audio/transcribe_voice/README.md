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
- **Small Model**: Used in quick mode for optimal speed/quality balance
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
├── transcribe_voice_launcher.py     # Unified launcher script
├── config.json                      # Configuration settings
├── requirements.txt                 # Python dependencies
├── setup.py                         # Package setup script
├── pyproject.toml                   # Modern Python packaging
├── quick_spanish.bat               # Quick Spanish launcher (batch)
├── quick_english.bat               # Quick English launcher (batch)
├── quick_spanish_streamdeck.bat    # Stream Deck Spanish launcher
├── quick_english_streamdeck.bat    # Stream Deck English launcher
├── quick_spanish.ps1               # Quick Spanish launcher (PowerShell)
├── quick_english.ps1               # Quick English launcher (PowerShell)
├── test_cuda.py                    # CUDA functionality test
├── test_cuda.bat                   # CUDA test batch file
├── test_cuda.ps1                   # CUDA test PowerShell script
├── CUDA.md                         # CUDA setup guide
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
- Follow CUDA setup guide below
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

## 🎯 CUDA Setup and GPU Acceleration

### Overview
This guide will help you activate CUDA on your Windows system with NVIDIA GeForce GTX 1070 GPU and configure it for use with Python virtual environments.

### Prerequisites
- NVIDIA GeForce GTX 1070 GPU (8GB VRAM)
- Windows 10/11 with PowerShell
- Python virtual environment at `E:\automation\automation\.venv`

### Step 1: Verify GPU and Drivers

Check if your NVIDIA GPU is detected:

```powershell
nvidia-smi
```

Expected output should show your GTX 1070 with driver information.

### Step 2: Install CUDA Toolkit

#### Download CUDA 12.1
1. Visit: https://developer.nvidia.com/cuda-12-1-0-download-archive
2. Select: Windows → Local → exe (local) → Download
3. Run the installer as Administrator
4. Choose "Custom" installation
5. Select these components:
   - CUDA Toolkit 12.1
   - CUDA Visual Studio Integration
   - CUDA Samples
   - CUDA Documentation

#### Verify CUDA Installation
```powershell
nvcc --version
```

Expected output:
```
nvcc: NVIDIA (R) Cuda compiler driver
Copyright (c) 2005-2023 NVIDIA Corporation
Built on Mon_Apr__3_17:36:15_PDT_2023
Cuda compilation tools, release 12.1, V12.1.105
Build cuda_12.1.r12.1/compiler.32688072_0
```

### Step 3: Install cuDNN (Optional but Recommended)

1. Download cuDNN from: https://developer.nvidia.com/cudnn
2. Sign in to NVIDIA Developer account (create if needed)
3. Download cuDNN for CUDA 12.x
4. Extract files to: `C:\Program Files\NVIDIA GPU Computing Toolkit\CUDA\v12.1\`

### Step 4: Configure Python Virtual Environment

#### Remove CPU-only PyTorch
```powershell
cd E:\automation\automation

# Remove current CPU-only PyTorch
& ".\.venv\Scripts\python.exe" -m pip uninstall torch torchvision torchaudio -y
```

#### Install CUDA-enabled PyTorch
```powershell
# Install PyTorch with CUDA 12.1 support
& ".\.venv\Scripts\python.exe" -m pip install torch torchvision torchaudio --index-url https://download.pytorch.org/whl/cu121
```

### Step 5: Verify CUDA Activation

#### Basic CUDA Test
```powershell
& ".\.venv\Scripts\python.exe" -c "
import torch
print('PyTorch version:', torch.__version__)
print('CUDA available:', torch.cuda.is_available())
print('CUDA version:', torch.version.cuda)
print('GPU count:', torch.cuda.device_count())
print('GPU name:', torch.cuda.get_device_name(0) if torch.cuda.is_available() else 'N/A')
"
```

Expected output:
```
PyTorch version: 2.8.0+cu121
CUDA available: True
CUDA version: 12.1
GPU count: 1
GPU name: NVIDIA GeForce GTX 1070
```

#### GPU Memory Test
```powershell
& ".\.venv\Scripts\python.exe" -c "
import torch
if torch.cuda.is_available():
    device = torch.device('cuda')
    x = torch.randn(1000, 1000).to(device)
    y = torch.matmul(x, x)
    print('✅ CUDA working! Matrix multiplication successful')
    print('GPU Memory used:', torch.cuda.memory_allocated(device) / 1024**2, 'MB')
else:
    print('❌ CUDA not available')
"
```

### Step 6: Environment Variables (Optional)

Add these to your system environment variables for better CUDA detection:

```powershell
# Add to System Environment Variables
CUDA_PATH = C:\Program Files\NVIDIA GPU Computing Toolkit\CUDA\v12.1
CUDA_HOME = C:\Program Files\NVIDIA GPU Computing Toolkit\CUDA\v12.1
Path += %CUDA_PATH%\bin
Path += %CUDA_PATH%\libnvvp
```

### CUDA Troubleshooting

#### Common Issues

1. **"CUDA not available" after installation**
   - Restart your computer
   - Check if virtual environment is activated correctly
   - Verify PyTorch installation: `pip list | findstr torch`

2. **Memory allocation errors**
   - Close other GPU-intensive applications
   - Check GPU memory usage: `nvidia-smi`

3. **Version compatibility issues**
   - Ensure CUDA toolkit version matches PyTorch CUDA version
   - GTX 1070 supports CUDA 12.x

#### Check GPU Memory Usage
```powershell
nvidia-smi --query-gpu=memory.used,memory.total --format=csv
```

### Performance Tips

1. **GPU Memory Management**
   ```python
   # Clear GPU cache
   torch.cuda.empty_cache()

   # Check memory usage
   print(torch.cuda.memory_summary())
   ```

2. **Data Transfer Optimization**
   ```python
   # Use pinned memory for faster CPU-GPU transfers
   data = torch.randn(1000, 1000).pin_memory()

   # Use non_blocking transfers
   tensor = tensor.to(device, non_blocking=True)
   ```

3. **Multi-GPU (if applicable)**
   ```python
   # Use DataParallel for multiple GPUs
   model = torch.nn.DataParallel(model)
   ```

### Testing Your CUDA Setup

#### Simple CUDA Benchmark
```python
import torch
import time

device = torch.device('cuda' if torch.cuda.is_available() else 'cpu')
print(f'Using device: {device}')

# Matrix multiplication benchmark
sizes = [1000, 2000, 5000]

for size in sizes:
    a = torch.randn(size, size).to(device)
    b = torch.randn(size, size).to(device)

    start_time = time.time()
    c = torch.matmul(a, b)
    torch.cuda.synchronize()  # Wait for GPU computation to complete
    end_time = time.time()

    print(f'Size {size}x{size}: {end_time - start_time:.4f} seconds')
```

### CUDA Files in This Directory

- **`CUDA.md`** - Complete CUDA installation and setup guide
- **`test_cuda.py`** - Python script to verify CUDA functionality
- **`test_cuda.bat`** - Windows batch file to run CUDA tests
- **`test_cuda.ps1`** - PowerShell script to run CUDA tests

#### Quick CUDA Test
```powershell
.\test_cuda.ps1
```
or
```batch
test_cuda.bat
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
2. Review CUDA setup guide above for GPU-related issues
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

*Last updated: September 2024*
*Tested on: Windows 11, NVIDIA GeForce GTX 1070, CUDA 12.1*