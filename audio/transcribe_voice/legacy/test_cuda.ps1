# CUDA Test Script Runner (PowerShell)
# Run this script to verify CUDA installation and functionality

Write-Host "CUDA Test Script Runner" -ForegroundColor Cyan
Write-Host "=======================" -ForegroundColor Cyan
Write-Host ""

# Set working directory
Set-Location "E:\automation\automation"

Write-Host "Running CUDA test with virtual environment..." -ForegroundColor Yellow
Write-Host ""

# Run the test script using the virtual environment
& ".\.venv\Scripts\python.exe" "audio\transcribe_voice\test_cuda.py"

Write-Host ""
Write-Host "Test completed. Press Enter to exit..." -ForegroundColor Green
Read-Host
