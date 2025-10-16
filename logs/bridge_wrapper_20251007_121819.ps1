$ErrorActionPreference='Continue'
$log = 'C:\AI\asagake\logs\bridge_20251007_121819.txt'
"START $(Get-Date -Format o)" | Out-File -FilePath $log -Encoding utf8
try { & 'py C:\AI\asagake\scripts\bridge_update_candidates.py' *>&1 | Tee-Object -FilePath $log -Append ; $code = $LASTEXITCODE }
catch { "EXCEPTION: $($_)" | Out-File -FilePath $log -Append ; $code = 1 }
"END $(Get-Date -Format o) exit=" + $code | Out-File -FilePath $log -Append
exit $code
