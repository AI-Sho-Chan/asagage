param(
  [string]$RepoRoot = "C:\\AI\\asagake"
)

$taskName = "Asagake-Nightly"
$action = "$RepoRoot\scripts\nightly_build.cmd"

try { schtasks /Delete /TN $taskName /F 2>$null | Out-Null } catch {}

# Run daily at 18:05, highest privileges not strictly required but helps reliability
schtasks /Create `
  /TN $taskName `
  /TR $action `
  /SC DAILY `
  /ST 18:05 `
  /RL HIGHEST `
  /F | Out-Null

Write-Host "Registered task $taskName to run daily at 18:05"

