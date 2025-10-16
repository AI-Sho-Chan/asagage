$ErrorActionPreference='Continue'
$log = 'C:\AI\asagake\logs\pipeline_2025-10-08_154254.txt'
"START $(Get-Date -Format o)" | Out-File -FilePath $log -Encoding utf8
try { & 'C:\AI\asagake\scripts\run_full.ps1' *>&1 | Tee-Object -FilePath $log -Append ; $code = $LASTEXITCODE }
catch { "EXCEPTION: $($_)" | Out-File -FilePath $log -Append ; $code = 1 }
"END $(Get-Date -Format o) exit=" + $code | Out-File -FilePath $log -Append
exit $code
