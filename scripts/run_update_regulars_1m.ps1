param(
  [string]$Repo = "C:\AI\asagake",
  [string]$Python = "C:\Python313\python.exe",
  [switch]$SkipTopVolBuild,
  [int]$UniverseTopN = 300,
  [int]$UniverseLookbackDays = 5,
  [int]$RegularsTopN = 200,
  [int]$RegularsLookbackFiles = 20,
  [int]$HistoryDays = 5,
  [int]$BackfillDays = 20,
  [int]$BatchSize = 120,
  [double]$PauseSeconds = 0.2
)

$ErrorActionPreference = "Stop"

Set-Location $Repo

$tag = Get-Date -Format "yyyyMMdd"
$logDir = Join-Path $Repo "logs"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
$logPath = Join-Path $logDir "update_regulars_1m_$tag.log"

function Run-Logged {
  param(
    [string[]]$PyArgs
  )

  $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = $Python
  $psi.Arguments = ($PyArgs -join " ")
  $psi.WorkingDirectory = $Repo
  $psi.RedirectStandardOutput = $true
  $psi.RedirectStandardError = $true
  $psi.UseShellExecute = $false

  $proc = New-Object System.Diagnostics.Process
  $proc.StartInfo = $psi
  $null = $proc.Start()
  $stdout = $proc.StandardOutput.ReadToEnd()
  $stderr = $proc.StandardError.ReadToEnd()
  $proc.WaitForExit()

  "[$ts] cmd: $Python $($PyArgs -join ' ')" | Out-File -FilePath $logPath -Append -Encoding utf8
  if ($stdout) { "[$ts] stdout`n$stdout" | Out-File -FilePath $logPath -Append -Encoding utf8 }
  if ($stderr) { "[$ts] stderr`n$stderr" | Out-File -FilePath $logPath -Append -Encoding utf8 }

  if ($proc.ExitCode -ne 0) {
    throw "Command failed (exit=$($proc.ExitCode)): $Python $($PyArgs -join ' ')"
  }
}

("[{0}] start update_regulars_1m" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss")) | Out-File -FilePath $logPath -Append -Encoding utf8

if (-not $SkipTopVolBuild) {
  Run-Logged @(
    "tools/build_master_topvol_universe.py",
    "--topn", $UniverseTopN,
    "--lookback", $UniverseLookbackDays,
    "--tag", $tag
  )
}

Run-Logged @(
  "tools/build_top_regulars_universe.py",
  "--lookback-files", $RegularsLookbackFiles,
  "--topn", $RegularsTopN,
  "--tag", $tag,
  "--update-ever"
)

$regulars = "data/universe/top_regulars_ever.csv"
if (-not (Test-Path $regulars)) {
  throw "Regulars CSV not found: $regulars"
}

Run-Logged @(
  "tools/update_minute_cache.py",
  "--codes-file", $regulars,
  "--history-days", $HistoryDays,
  "--backfill-days", $BackfillDays,
  "--batch-size", $BatchSize,
  "--pause", $PauseSeconds
)

("[{0}] done update_regulars_1m" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss")) | Out-File -FilePath $logPath -Append -Encoding utf8
