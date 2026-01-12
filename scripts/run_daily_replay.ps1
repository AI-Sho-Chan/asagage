param(
  [string]$BaseDir = "C:\AI\asagake",
  [string]$PythonExe = "C:\Python313\python.exe",
  [string]$DateTag = "",
  [string]$Recipient = "shouichi.ikeda@gmail.com"
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
Add-Content -Path $logPath -Value ("[{0}] === DailyReplay start (date={1}) ===" -f $ts, $DateTag) -Encoding utf8

$replayScript = Join-Path $BaseDir "tools\simulate_daily_replay.py"
if (-not (Test-Path $replayScript)) {
  Add-Content -Path $logPath -Value ("[{0}] [error] replay script not found: {1}" -f (Get-Date).ToString("s"), $replayScript) -Encoding utf8
  exit 1
}

try {
  & $PythonExe $replayScript --date $DateTag --email --recipient $Recipient 2>&1 | Out-File -FilePath $logPath -Append -Encoding utf8
  $exitCode = $LASTEXITCODE
  Add-Content -Path $logPath -Value ("[{0}] exit_code={1}" -f (Get-Date).ToString("s"), $exitCode) -Encoding utf8
  Add-Content -Path $logPath -Value ("[{0}] === DailyReplay end ===" -f (Get-Date).ToString("s")) -Encoding utf8
  if ($exitCode -eq 0) { exit 0 } else { exit 1 }
} catch {
  Add-Content -Path $logPath -Value ("[{0}] [error] {1}" -f (Get-Date).ToString("s"), $_.Exception.ToString()) -Encoding utf8
  Add-Content -Path $logPath -Value ("[{0}] === DailyReplay end ===" -f (Get-Date).ToString("s")) -Encoding utf8
  exit 1
}
