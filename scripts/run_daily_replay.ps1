param(
  [string]$DateTag = (Get-Date -Format "yyyyMMdd"),
  [double]$Nominal = 10000000,
  [int]$MaxAttempts = 12,
  [int]$RetrySleepSeconds = 600
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

$summaryPath = Join-Path $repo ("analysis/daily_replay_{0}.json" -f $DateTag)
$reportPath = Join-Path $repo ("analysis/daily_replay_{0}_mail.txt" -f $DateTag)

Write-Log ("=== DailyReplay start (date={0}) ===" -f $DateTag)

$simulateOk = $false
$lastError = ""

for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
  try {
    Write-Log ("simulate attempt {0}/{1}" -f $attempt, $MaxAttempts)

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

    $simulateOk = $true
    break
  } catch {
    $lastError = $_.Exception.Message
    Write-Log ("[warn] simulate failed: {0}" -f $lastError)
    if ($attempt -lt $MaxAttempts) {
      Write-Log ("sleep {0}s then retry" -f $RetrySleepSeconds)
      Start-Sleep -Seconds $RetrySleepSeconds
    }
  }
}

if (-not $simulateOk) {
  Write-Log ("[error] simulate did not succeed after {0} attempts; writing placeholder summary" -f $MaxAttempts)
  $placeholder = @{
    date = $DateTag
    status = "NOT_READY"
    attempts = $MaxAttempts
    note = "Yahoo 1m data may not be ready yet; rerun later."
    last_error = $lastError
  } | ConvertTo-Json -Depth 4
  $placeholder | Out-File -FilePath $summaryPath -Encoding utf8

  $placeholderText = @"
ASAGAKE DailyReplay $DateTag

Status: NOT_READY
Note: Yahoo 1m data may not be ready yet; rerun later.
LastError: $lastError
"@
  $placeholderText | Out-File -FilePath $reportPath -Encoding utf8
}

try {
  $smtpPath = Join-Path $repo "state/smtp.json"
  if ((Test-Path $summaryPath) -and (Test-Path $smtpPath)) {
    $summary = Get-Content $summaryPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $smtpConf = Get-Content $smtpPath -Raw -Encoding UTF8 | ConvertFrom-Json

    $cred = New-Object System.Management.Automation.PSCredential(
      $smtpConf.user,
      (ConvertTo-SecureString $smtpConf.pass -AsPlainText -Force)
    )

    if (Test-Path $reportPath) {
      $body = Get-Content $reportPath -Raw -Encoding UTF8
    } else {
      $body = Get-Content $summaryPath -Raw -Encoding UTF8
    }

    $subject = ("ASAGAKE DailyReplay {0}" -f $DateTag)
    if ($summary.PSObject.Properties.Name -contains "status" -and $summary.status -eq "NOT_READY") {
      $subject = ("ASAGAKE DailyReplay {0} (NOT_READY)" -f $DateTag)
    }

    $mailParams = @{
      To         = "shouichi.ikeda@gmail.com"
      From       = $smtpConf.user
      Subject    = $subject
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

# Always exit 0 so the scheduled task does not become a recurring failure when Yahoo data is delayed.
exit 0
