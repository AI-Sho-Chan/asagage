param(
  [string]$TaskName = "Asagake-Update1m-Regulars",
  [string]$Time = "18:40",
  [string]$Days = "MON,TUE,WED,THU"
)

$scriptPath = "C:\AI\asagake\scripts\run_update_regulars_1m.ps1"
if (-not (Test-Path $scriptPath)) {
  throw "Script not found: $scriptPath"
}

try {
  schtasks /Delete /TN $TaskName /F 2>$null | Out-Null
} catch {}

$tr = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""

schtasks /Create `
  /TN $TaskName `
  /TR $tr `
  /SC WEEKLY `
  /D $Days `
  /ST $Time `
  /RL LIMITED `
  /F | Out-Null

Write-Host "Registered task '$TaskName' at $Time ($Days)."

