$taskName = "Asagake-Update1m"
$scriptPath = "C:\AI\asagake\scripts\run_update_all_1m.cmd"

try {
  schtasks /Delete /TN $taskName /F 2>$null | Out-Null
} catch {}

schtasks /Create `
  /TN $taskName `
  /TR "`"$scriptPath`"" `
  /SC WEEKLY `
  /D MON,TUE,WED,THU,FRI `
  /ST 05:30 `
  /RL LIMITED `
  /F | Out-Null

Write-Host "Registered task '$taskName' to run daily at 05:30 (MON-FRI)"
