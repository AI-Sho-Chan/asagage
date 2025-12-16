param(
  [string]$DateTag = (Get-Date -Format "yyyyMMdd"),
  [double]$Nominal = 10000000
)

$ErrorActionPreference = "Stop"

$repo = "C:/AI/asagake"
$python = "C:/Python313/python.exe"

Set-Location $repo

if (-not $DateTag) {
  $DateTag = (Get-Date -Format "yyyyMMdd")
}

$logDir = Join-Path $repo "logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$logPath = Join-Path $logDir ("daily_replay_task_{0}.log" -f $DateTag)

function Write-Log {
  param([string]$Message)
  $ts = Get-Date -Format s
  ("[{0}] {1}" -f $ts, $Message) | Out-File -FilePath $logPath -Append -Encoding utf8
}

Write-Log ("=== DailyReplay start (date={0}) ===" -f $DateTag)

try {
  $pyArgs = @(
    "tools/simulate_daily_replay.py",
    "--date", $DateTag,
    "--nominal", [string][int][math]::Round($Nominal)
  )

  $out = & $python @pyArgs 2>&1
  $out | Out-File -FilePath $logPath -Append -Encoding utf8
  if ($LASTEXITCODE -ne 0) {
    throw "simulate_daily_replay.py failed (exit=$LASTEXITCODE)"
  }
} catch {
  Write-Log ("[error] simulate failed: {0}" -f $_.Exception.Message)
  throw
}

# If `analysis/daily_replay_<date>.json` and `state/smtp.json` exist, send summary email.
try {
  $summaryPath = Join-Path $repo ("analysis/daily_replay_{0}.json" -f $DateTag)
  $smtpPath = Join-Path $repo "state/smtp.json"

  if ((Test-Path $summaryPath) -and (Test-Path $smtpPath)) {
    $summary = Get-Content $summaryPath -Raw | ConvertFrom-Json
    $smtpConf = Get-Content $smtpPath -Raw | ConvertFrom-Json

    $cred = New-Object System.Management.Automation.PSCredential(
      $smtpConf.user,
      (ConvertTo-SecureString $smtpConf.pass -AsPlainText -Force)
    )

    $body = $summary | ConvertTo-Json -Depth 8

    $mailParams = @{
      To         = "shouichi.ikeda@gmail.com"
      From       = $smtpConf.user
      Subject    = ("ASAGAKE DailyReplay {0}" -f $DateTag)
      Body       = $body
      SmtpServer = $smtpConf.host
      Port       = [int]$smtpConf.port
      UseSsl     = $true
      Credential = $cred
    }

    Send-MailMessage @mailParams
    Write-Log "Mail sent to shouichi.ikeda@gmail.com"
  } else {
    Write-Log "No smtp.json or summary not found; skip mail."
  }
} catch {
  Write-Log ("Mail send failed: {0}" -f $_.Exception.Message)
}

Write-Log "=== DailyReplay end ==="
