# CUDA Setup for Voice Transcription

This directory contains CUDA setup files for enabling GPU acceleration in voice transcription projects.

## Files

- **`CUDA.md`** - Complete CUDA installation and setup guide
- **`test_cuda.py`** - Python script to verify CUDA functionality
- **`test_cuda.bat`** - Windows batch file to run CUDA tests
- **`test_cuda.ps1`** - PowerShell script to run CUDA tests
- **`README_CUDA.md`** - This file

## Quick Start

1. **Install CUDA following `CUDA.md`**
2. **Test your setup:**
   ```powershell
   .\test_cuda.ps1
   ```
   or
   ```batch
   test_cuda.bat
   ```

## Usage with Voice Transcription

Once CUDA is activated, you can use GPU acceleration for:
- Audio processing with PyTorch
- Machine learning models for speech recognition
- Faster transcription processing
- GPU-accelerated audio analysis

## Requirements

- NVIDIA GeForce GTX 1070 (or compatible GPU)
- Windows 10/11
- Python virtual environment at `E:\automation\automation\.venv`

## Support

See `CUDA.md` for detailed installation instructions and troubleshooting.
