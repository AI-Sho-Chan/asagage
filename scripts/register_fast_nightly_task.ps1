param(
  [string]$Time = '16:30',
  [string]$TaskName = 'Asagake Nightly FAST'
)
$ErrorActionPreference='Stop'
$script = 'C:\AI\asagake\scripts\run_fast_nightly.ps1'
if (!(Test-Path $script)) { throw "Script not found: $script" }
$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$script`""
$trigger = New-ScheduledTaskTrigger -Daily -At $Time
$settings = New-ScheduledTaskSettingsSet -Compatibility Win8 -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
try {
  Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
} catch {}
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Description 'Run FAST nightly (Top150, ASHA+Bayes, H1/H3)' | Out-Null
Write-Output "registered=$TaskName at $Time"
