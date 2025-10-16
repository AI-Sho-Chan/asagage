$ErrorActionPreference='Continue'
$log = 'C:\AI\asagake\logs\retune_2025-10-08_124347.txt'
"START $(Get-Date -Format o)" | Out-File -FilePath $log -Encoding utf8
try { & 'py C:\AI\asagake\scripts\optimize_thresholds.py' *>&1 | Tee-Object -FilePath $log -Append ; $code = $LASTEXITCODE }
catch { "EXCEPTION: $($_)" | Out-File -FilePath $log -Append ; $code = 1 }
"END $(Get-Date -Format o) exit=" + $code | Out-File -FilePath $log -Append
exit $code
