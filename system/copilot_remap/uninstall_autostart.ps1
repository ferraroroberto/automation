# Removes the RemapCopilotKey scheduled task and stops any running instance.

$ErrorActionPreference = "Stop"
$taskName = "RemapCopilotKey"

if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Host "Removed scheduled task '$taskName'." -ForegroundColor Green
} else {
    Write-Host "No task '$taskName' registered. Nothing to do."
}

# Kill any lingering pythonw.exe running our script.
$here   = Split-Path -Parent $MyInvocation.MyCommand.Path
$script = Join-Path $here "copilot_remap.py"
Get-CimInstance Win32_Process -Filter "Name='pythonw.exe'" |
    Where-Object { $_.CommandLine -like "*copilot_remap.py*" } |
    ForEach-Object {
        Write-Host "Stopping pythonw.exe pid=$($_.ProcessId)"
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
    }
