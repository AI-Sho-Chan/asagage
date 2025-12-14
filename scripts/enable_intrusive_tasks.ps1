param(
  [string[]]$TaskNames = @(
    "ASAGAKE_BoardLogger",
    "Asagake-BoardLogger",
    "Asagake-Update1m",
    "Asagake-Update1m-Today"
  )
)

foreach ($name in $TaskNames) {
  schtasks /Query /TN $name 2>$null | Out-Null
  if ($LASTEXITCODE -ne 0) {
    Write-Output "not_found=$name"
    continue
  }
  schtasks /Change /TN $name /Enable | Out-Null
  Write-Output "enabled=$name"
}

