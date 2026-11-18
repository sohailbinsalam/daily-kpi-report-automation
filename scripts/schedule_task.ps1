# Registers a Windows Task Scheduler job that runs the pipeline every morning.
# Usage (from the project folder):
#   powershell -ExecutionPolicy Bypass -File scripts\schedule_task.ps1 -Time 06:30
param(
    [string]$Time = "06:30",
    [string]$TaskName = "Daily KPI Report"
)

$project = Split-Path -Parent $PSScriptRoot
$python  = Join-Path $project ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) { $python = (Get-Command python).Source }

$action  = New-ScheduledTaskAction -Execute $python -Argument "run_daily.py" -WorkingDirectory $project
$trigger = New-ScheduledTaskTrigger -Daily -At $Time
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 30)

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
    -Description "Excel KPI file -> ETL -> PowerPoint -> email" -Force | Out-Null

Write-Host "Scheduled '$TaskName' every day at $Time using $python"
