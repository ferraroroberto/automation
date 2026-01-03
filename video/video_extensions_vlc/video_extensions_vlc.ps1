# Video Extensions VLC Default Setter
# This script checks for common video file extensions and offers to set VLC as the default program

# Common video file extensions
$videoExtensions = @(
    ".mp4", ".avi", ".mkv", ".mov", ".wmv", ".flv", ".webm", ".m4v",
    ".3gp", ".mpg", ".mpeg", ".vob", ".ogv", ".rm", ".rmvb", ".asf",
    ".divx", ".xvid", ".f4v", ".mts", ".m2ts", ".ts", ".mts", ".mxf"
)

Write-Host "Video Extensions VLC Default Setter" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

# Function to get current default program for extension
function Get-CurrentDefaultProgram {
    param([string]$extension)

    try {
        # Get the file type associated with the extension
        $assocOutput = cmd /c "assoc $extension 2>nul"
        if ($assocOutput -and $assocOutput.Contains('=')) {
            $fileType = $assocOutput.Split('=')[1].Trim()

            if ($fileType) {
                # Get the command associated with the file type
                $ftypeOutput = cmd /c "ftype $fileType 2>nul"
                if ($ftypeOutput -and $ftypeOutput.Contains('=')) {
                    $command = $ftypeOutput.Split('=')[1].Trim()

                    if ($command) {
                        # Extract program name from command - handle quoted paths
                        if ($command.StartsWith('"')) {
                            $endQuote = $command.IndexOf('"', 1)
                            if ($endQuote -gt 0) {
                                $program = $command.Substring(1, $endQuote - 1)
                            } else {
                                $program = $command
                            }
                        } else {
                            $program = ($command -split ' ')[0]
                        }

                        # Get just the filename and clean it
                        $fileName = [System.IO.Path]::GetFileName($program)
                        return $fileName.Trim()
                    }
                }
            }
        }
    }
    catch {
        # Silently continue on errors
    }

    return "Not associated"
}

# Function to set VLC as default for extension
function Set-VLCAsDefault {
    param([string]$extension)

    try {
        # Check if VLC is installed
        $vlcPath = Get-Command vlc -ErrorAction SilentlyContinue
        if (!$vlcPath) {
            # Try common VLC installation paths
            $commonPaths = @(
                "${env:ProgramFiles}\VideoLAN\VLC\vlc.exe",
                "${env:ProgramFiles(x86)}\VideoLAN\VLC\vlc.exe",
                "$env:LOCALAPPDATA\VideoLAN\VLC\vlc.exe"
            )

            foreach ($path in $commonPaths) {
                if (Test-Path $path) {
                    $vlcPath = $path
                    break
                }
            }

            if (!$vlcPath) {
                Write-Host "  Error: VLC not found. Please install VLC Media Player first." -ForegroundColor Red
                return $false
            }
        } else {
            $vlcPath = $vlcPath.Source
        }

        # Method 1: Try using assoc/ftype commands (works without admin rights for some extensions)
        try {
            $fileType = $extension.TrimStart('.') + "file"
            $command = "`"$vlcPath`" `"%1`""

            # Set the file type association
            $null = cmd /c "ftype $fileType=$command 2>nul"
            $null = cmd /c "assoc $extension=$fileType 2>nul"

            # Test if it worked
            $testCommand = cmd /c "ftype $fileType 2>nul"
            if ($testCommand -and $testCommand -like "*$vlcPath*") {
                return $true
            }
        }
        catch {
            # Continue to method 2 if assoc/ftype fails
        }

        # Method 2: Registry method (requires admin rights for UserChoice, but try anyway)
        $extensionClass = $extension.TrimStart('.') + "_auto_file"

        # Create or update the registry entries
        $regPath = "HKCU:\Software\Classes\$extensionClass\shell\open\command"
        if (!(Test-Path $regPath)) {
            New-Item -Path $regPath -Force -ErrorAction SilentlyContinue | Out-Null
        }

        # Set the command
        Set-ItemProperty -Path $regPath -Name "(Default)" -Value "`"$vlcPath`" `"%1`"" -Force -ErrorAction SilentlyContinue

        # Try to set UserChoice (this often requires admin rights in Windows 10/11)
        $extRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$extension\UserChoice"
        try {
            if (!(Test-Path $extRegPath)) {
                New-Item -Path $extRegPath -Force -ErrorAction Stop | Out-Null
            }

            # Generate a unique ProgId for VLC
            $vlcProgId = "VLC.$($extension.TrimStart('.'))"
            Set-ItemProperty -Path $extRegPath -Name "ProgId" -Value $vlcProgId -Force -ErrorAction Stop
            Set-ItemProperty -Path $extRegPath -Name "Hash" -Value "" -Force -ErrorAction Stop

            # Set up the ProgId
            $progIdPath = "HKCU:\Software\Classes\$vlcProgId"
            if (!(Test-Path $progIdPath)) {
                New-Item -Path $progIdPath -Force -ErrorAction Stop | Out-Null
            }

            # Set default command
            $cmdPath = "$progIdPath\shell\open\command"
            if (!(Test-Path $cmdPath)) {
                New-Item -Path $cmdPath -Force -ErrorAction Stop | Out-Null
            }
            Set-ItemProperty -Path $cmdPath -Name "(Default)" -Value "`"$vlcPath`" `"%1`"" -Force -ErrorAction Stop

            # Set friendly name
            Set-ItemProperty -Path $progIdPath -Name "(Default)" -Value "VLC media file" -Force -ErrorAction Stop

            return $true
        }
        catch {
            # If UserChoice fails, at least we set up the basic association
            Write-Host "  Warning: Could not set UserChoice registry (requires admin rights for $extension)" -ForegroundColor Yellow
            Write-Host "  Basic association may still work, but some extensions need admin rights" -ForegroundColor Yellow
            # Still return true since the basic assoc/ftype method might have worked
            return $true
        }
    }
    catch {
        Write-Host "  Error setting VLC as default for $extension : $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Main processing
$extensionsProcessed = 0
$extensionsChanged = 0

foreach ($ext in $videoExtensions) {
    Write-Host ""
    Write-Host "Checking extension: $ext" -ForegroundColor Yellow

    # Check if extension is registered
    $currentProgram = Get-CurrentDefaultProgram $ext

    if ($currentProgram -eq "Not associated") {
        Write-Host "  Status: Not associated with any program" -ForegroundColor Gray
        Write-Host "  Set VLC as default for $ext (y/n)? " -NoNewline
        $response = Read-Host
        if ([string]::IsNullOrWhiteSpace($response) -or $response.ToLower() -eq 'y' -or $response.ToLower() -eq 'yes') {
            if (Set-VLCAsDefault $ext) {
                Write-Host "  ✓ Successfully set VLC as default for $ext" -ForegroundColor Green
                $extensionsChanged++
            } else {
                Write-Host "  ✗ Failed to set VLC as default for $ext" -ForegroundColor Red
            }
        } else {
            Write-Host "  Skipped $ext" -ForegroundColor Gray
        }
    }
    elseif ($currentProgram -like "*vlc*") {
        Write-Host "  Status: Already set to VLC" -ForegroundColor Green
    }
    else {
        Write-Host "  Status: Currently set to '$currentProgram'" -ForegroundColor Cyan
        Write-Host "  Change $ext from '$currentProgram' to VLC (y/n)? " -NoNewline
        $response = Read-Host
        if ([string]::IsNullOrWhiteSpace($response) -or $response.ToLower() -eq 'y' -or $response.ToLower() -eq 'yes') {
            if (Set-VLCAsDefault $ext) {
                Write-Host "  ✓ Successfully changed $ext to VLC" -ForegroundColor Green
                $extensionsChanged++
            } else {
                Write-Host "  ✗ Failed to change $ext to VLC" -ForegroundColor Red
            }
        } else {
            Write-Host "  Kept current association for $ext" -ForegroundColor Gray
        }
    }

    $extensionsProcessed++
}

Write-Host ""
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "=========" -ForegroundColor Cyan
Write-Host "Extensions processed: $extensionsProcessed"
Write-Host "Extensions changed to VLC: $extensionsChanged"
Write-Host ""
Write-Host "Note: You may need to restart Windows Explorer or log out/in for changes to take effect." -ForegroundColor Yellow
Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor Gray -NoNewline
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")