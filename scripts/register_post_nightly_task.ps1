$taskName = "Asagake-PostNightly"
$workDir = "C:\AI\asagake"
$actionCmd = "C:\Windows\System32\cmd.exe"
$actionArgs = "/c scripts\run_post_nightly.cmd"

$action = New-ScheduledTaskAction -Execute $actionCmd -Argument $actionArgs -WorkingDirectory $workDir
$trigger = New-ScheduledTaskTrigger -Daily -At 7:20AM
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

try {
  Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
} catch {}

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Description "Run Asagake post-nightly tasks (logs export + size plan)" | Out-Null
Write-Host "Registered task '$taskName' to run daily at 07:20"

