param(
  [string]$BaseDir = "C:\AI\asagake",
  [string]$PythonExe = "",
  [string]$DateTag = "",
  [string]$Recipient = "shouichi.ikeda@gmail.com"
)

$ErrorActionPreference = "Stop"

Set-Location $BaseDir

if ([string]::IsNullOrWhiteSpace($PythonExe)) {
  $candidateVenv = Join-Path $BaseDir ".venv\Scripts\python.exe"
  if (Test-Path $candidateVenv) {
    $PythonExe = $candidateVenv
  } else {
    $PythonExe = "C:\Python313\python.exe"
  }
}

if (-not (Test-Path $PythonExe)) {
  throw "Python not found: $PythonExe (set -PythonExe or create .venv)"
}

New-Item -ItemType Directory -Force -Path (Join-Path $BaseDir "logs") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $BaseDir "analysis") | Out-Null

if ([string]::IsNullOrWhiteSpace($DateTag)) {
  $DateTag = (Get-Date).ToString("yyyyMMdd")
}

$logPath = Join-Path $BaseDir ("logs\daily_replay_task_{0}.log" -f $DateTag)
$ts = (Get-Date).ToString("s")
Add-Content -Path $logPath -Value ("[{0}] === DailyReplay start (date={1}) ===" -f $ts, $DateTag) -Encoding utf8
Add-Content -Path $logPath -Value ("[{0}] whoami={1}" -f (Get-Date).ToString("s"), [System.Security.Principal.WindowsIdentity]::GetCurrent().Name) -Encoding utf8

$replayScript = Join-Path $BaseDir "tools\simulate_daily_replay.py"
if (-not (Test-Path $replayScript)) {
  Add-Content -Path $logPath -Value ("[{0}] [error] replay script not found: {1}" -f (Get-Date).ToString("s"), $replayScript) -Encoding utf8
  exit 1
}

$sentFlagPath = Join-Path $BaseDir ("logs\daily_replay_sent_{0}.flag" -f $DateTag)
if (Test-Path $sentFlagPath) {
  Add-Content -Path $logPath -Value ("[{0}] [warn] already sent today; skip run (flag={1})" -f (Get-Date).ToString("s"), $sentFlagPath) -Encoding utf8
  Add-Content -Path $logPath -Value ("[{0}] === DailyReplay end ===" -f (Get-Date).ToString("s")) -Encoding utf8
  exit 0
}

$lockPath = Join-Path $BaseDir ("logs\daily_replay_lock_{0}.lock" -f $DateTag)
$lockStream = $null
try {
  if (Test-Path $lockPath) {
    $age = (Get-Date) - (Get-Item $lockPath).LastWriteTime
    if ($age.TotalHours -gt 6) {
      Remove-Item -Force $lockPath -ErrorAction SilentlyContinue
    }
  }

  $lockStream = [System.IO.File]::Open($lockPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
  $lockBytes = [System.Text.Encoding]::UTF8.GetBytes(("pid={0} started={1}`n" -f $PID, (Get-Date).ToString("s")))
  $lockStream.Write($lockBytes, 0, $lockBytes.Length)
  $lockStream.Flush()
} catch {
  Add-Content -Path $logPath -Value ("[{0}] [warn] lock exists; skip run (lock={1})" -f (Get-Date).ToString("s"), $lockPath) -Encoding utf8
  exit 0
}

try {
  $output = & $PythonExe $replayScript --date $DateTag --email --recipient $Recipient 2>&1
  $output | Out-File -FilePath $logPath -Append -Encoding utf8
  $exitCode = $LASTEXITCODE
  Add-Content -Path $logPath -Value ("[{0}] exit_code={1}" -f (Get-Date).ToString("s"), $exitCode) -Encoding utf8
  Add-Content -Path $logPath -Value ("[{0}] === DailyReplay end ===" -f (Get-Date).ToString("s")) -Encoding utf8
  exit $exitCode
} catch {
  $err = $_ | Out-String
  Add-Content -Path $logPath -Value ("[{0}] [error] {1}" -f (Get-Date).ToString("s"), $err.Trim()) -Encoding utf8
  Add-Content -Path $logPath -Value ("[{0}] === DailyReplay end ===" -f (Get-Date).ToString("s")) -Encoding utf8
  exit 1
} finally {
  try { if ($lockStream) { $lockStream.Close() } } catch {}
  try { Remove-Item -Force $lockPath -ErrorAction SilentlyContinue } catch {}
}
