param(
  [string]$PythonExe = "python",
  [string]$RepoRoot = "C:\\AI\\asagake"
)

$taskName = "Asagake-WeekendFullScan"
$action = "$RepoRoot\scripts\run_weekend_fullscan.cmd"

try { schtasks /Delete /TN $taskName /F 2>$null | Out-Null } catch {}

schtasks /Create `
  /TN $taskName `
  /TR $action `
  /SC WEEKLY `
  /D SAT `
  /ST 01:00 `
  /RL HIGHEST `
  /F | Out-Null

Write-Host "Registered task $taskName to run Saturdays at 01:00"
