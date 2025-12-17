param(
  [string]$Repo = "C:\AI\asagake",
  [string]$Python = "C:\Python313\python.exe",
  [string]$Gsutil = "",
  [string]$RegularsGcsBucket = "gs://asagage-weekend-output/yahoo_1m_regulars",
  [switch]$SkipRegularsMirror,
  [switch]$SkipRegularsUpload,
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

$resolvedGsutil = $Gsutil
if (-not $resolvedGsutil) {
  $gsutilCmd = Get-Command "gsutil.cmd" -ErrorAction SilentlyContinue
  if ($gsutilCmd) {
    $resolvedGsutil = $gsutilCmd.Source
  } else {
    $gsutilExe = Get-Command "gsutil" -ErrorAction SilentlyContinue
    if ($gsutilExe) {
      $resolvedGsutil = $gsutilExe.Source
    } else {
      $resolvedGsutil = "gsutil"
    }
  }
}

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

if (-not $SkipRegularsMirror) {
  $mirrorRoot = Join-Path $Repo "cache\\yahoo_1m_regulars"
  if (-not (Test-Path $mirrorRoot)) { New-Item -ItemType Directory -Force -Path $mirrorRoot | Out-Null }

  $codes = Import-Csv $regulars | ForEach-Object { $_.code } | Where-Object { $_ -and $_.Trim() } | Sort-Object -Unique
  foreach ($code in $codes) {
    $src = Join-Path $Repo ("data\\raw\\yahoo_1m\\{0}" -f $code)
    $dst = Join-Path $mirrorRoot $code
    if (-not (Test-Path $src)) { continue }
    if (-not (Test-Path $dst)) { New-Item -ItemType Directory -Force -Path $dst | Out-Null }
    # Mirror only parquet files; robocopy is incremental and fast when nothing changed.
    robocopy $src $dst *.parquet /MIR /R:2 /W:1 /NFL /NDL /NP /NJH /NJS | Out-Null
  }
}

if (-not $SkipRegularsUpload) {
  try {
    $mirrorRoot = Join-Path $Repo "cache\\yahoo_1m_regulars"
    if (-not (Test-Path $mirrorRoot)) {
      throw "Mirror dir not found: $mirrorRoot (run without -SkipRegularsMirror)"
    }

    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$ts] upload regulars to $RegularsGcsBucket" | Out-File -FilePath $logPath -Append -Encoding utf8

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $resolvedGsutil
    $psi.Arguments = ('-m rsync -r "{0}" "{1}"' -f $mirrorRoot, $RegularsGcsBucket)
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

    if ($stdout) { "[$ts] gsutil stdout`n$stdout" | Out-File -FilePath $logPath -Append -Encoding utf8 }
    if ($stderr) { "[$ts] gsutil stderr`n$stderr" | Out-File -FilePath $logPath -Append -Encoding utf8 }
    if ($proc.ExitCode -ne 0) {
      throw "gsutil rsync failed (exit=$($proc.ExitCode))"
    }
  } catch {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$ts] [warn] regulars upload skipped/failed: $($_.Exception.Message)" | Out-File -FilePath $logPath -Append -Encoding utf8
  }
}

("[{0}] done update_regulars_1m" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss")) | Out-File -FilePath $logPath -Append -Encoding utf8
