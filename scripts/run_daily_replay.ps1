param(
  [string]$BaseDir = "C:\AI\asagake",
  [string]$PythonExe = "C:\Python313\python.exe",
  [string]$DateTag = ""
)

$ErrorActionPreference = "Stop"

Set-Location $BaseDir

if (-not (Test-Path $PythonExe)) {
  throw "Python not found: $PythonExe"
}

New-Item -ItemType Directory -Force -Path (Join-Path $BaseDir "logs") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $BaseDir "analysis") | Out-Null

if ([string]::IsNullOrWhiteSpace($DateTag)) {
  $DateTag = (Get-Date).ToString("yyyyMMdd")
}

$logPath = Join-Path $BaseDir ("logs\daily_replay_task_{0}.log" -f $DateTag)
$ts = (Get-Date).ToString("s")
Add-Content -Path $logPath -Value ("[{0}] === DailyReplay start (date={1}) ===" -f $ts, $DateTag)

try {
  & $PythonExe tools\simulate_daily_replay.py --date-tag $DateTag 2>&1 | Tee-Object -FilePath $logPath -Append | Out-Null
  Add-Content -Path $logPath -Value ("[{0}] === DailyReplay end ===" -f (Get-Date).ToString("s"))
  exit 0
} catch {
  Add-Content -Path $logPath -Value ("[{0}] [error] {1}" -f (Get-Date).ToString("s"), $_.Exception.Message)
  exit 1
}

