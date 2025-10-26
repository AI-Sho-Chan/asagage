$taskName = "Asagake-BoardLogger"
$scriptPath = "C:\AI\asagake\scripts\run_board_logger_daemon.cmd"

try {
  schtasks /Delete /TN $taskName /F 2>$null | Out-Null
} catch {}

schtasks /Create `
  /TN $taskName `
  /TR "`"$scriptPath`"" `
  /SC WEEKLY `
  /D MON,TUE,WED,THU,FRI `
  /ST 08:55 `
  /RL LIMITED `
  /F | Out-Null

Write-Host "Registered task '$taskName' to run daily at 08:55 (MON-FRI)"
