# Quick English Transcription Launcher
# Launches transcription in quick mode without GUI

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

# Calculate venv path from script location
$venvPath = Join-Path $scriptDir "..\..\.venv"
$venvPath = Resolve-Path $venvPath

& "$venvPath\Scripts\python.exe" transcribe_voice_gui.py --quick-english
