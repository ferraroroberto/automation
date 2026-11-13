# Registers a per-user Scheduled Task that runs the Copilot key remapper at logon.
# Hidden window (uses pythonw.exe), no admin required. Re-run to refresh.

$ErrorActionPreference = "Stop"

$here     = Split-Path -Parent $MyInvocation.MyCommand.Path
$script   = Join-Path $here "copilot_remap.py"
$config   = Join-Path $here "chord.json"
$pythonw  = Join-Path (Split-Path -Parent (Split-Path -Parent $here)) ".venv\Scripts\pythonw.exe"
$taskName = "RemapCopilotKey"

if (-not (Test-Path $pythonw)) {
    Write-Error "pythonw.exe not found at $pythonw. Edit this script if your venv lives elsewhere."
}
if (-not (Test-Path $script)) {
    Write-Error "Script not found at $script."
}
if (-not (Test-Path $config)) {
    Write-Warning "No chord.json yet - run 'python copilot_remap.py detect' first, then re-run this installer."
}

$action    = New-ScheduledTaskAction -Execute $pythonw -Argument "`"$script`" run" -WorkingDirectory $here
$trigger   = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$settings  = New-ScheduledTaskSettingsSet `
                -AllowStartIfOnBatteries `
                -DontStopIfGoingOnBatteries `
                -StartWhenAvailable `
                -ExecutionTimeLimit (New-TimeSpan -Seconds 0) `
                -MultipleInstances IgnoreNew
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

$task = New-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -Principal $principal `
                          -Description "Remap Windows Copilot key chord back to Ctrl."

Register-ScheduledTask -TaskName $taskName -InputObject $task -Force | Out-Null

Write-Host "Registered scheduled task '$taskName'." -ForegroundColor Green
Write-Host "  Runs at logon for user: $env:USERNAME"
Write-Host "  Command: $pythonw `"$script`" run"
Write-Host ""
Write-Host "Start it now:  Start-ScheduledTask -TaskName $taskName"
Write-Host "Stop it now:   Stop-ScheduledTask  -TaskName $taskName"
Write-Host "Check status:  Get-ScheduledTask   -TaskName $taskName | Get-ScheduledTaskInfo"
