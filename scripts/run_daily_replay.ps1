param(
  [string]$DateTag = (Get-Date -Format 'yyyyMMdd'),
  [double]$Nominal = 10000000
)

$ErrorActionPreference = "Stop"
$repo = "C:/AI/asagake"
$python = "C:/Python313/python.exe"

Set-Location $repo

if (-not $DateTag) {
  $DateTag = (Get-Date -Format 'yyyyMMdd')
}

$logDir = Join-Path $repo "logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$logPath = Join-Path $logDir ("daily_replay_task_{0}.log" -f $DateTag)

("=== DailyReplay start {0} ===" -f (Get-Date -Format s)) | Out-File -FilePath $logPath -Append -Encoding utf8

$args = @(
  "tools/simulate_daily_replay.py",
  "--date", $DateTag,
  "--nominal", [string][int][math]::Round($Nominal)
)

try {
  & $python $args 2>&1 | Tee-Object -FilePath $logPath -Append
} catch {
  ("[error] simulate failed: {0}" -f $_.Exception.Message) | Out-File -FilePath $logPath -Append -Encoding utf8
  throw
}

# メール送信（state/smtp.json がある場合のみ）
try {
  $summaryPath = Join-Path $repo ("analysis/daily_replay_{0}.json" -f $DateTag)
  $smtpPath = Join-Path $repo "state/smtp.json"
  if ((Test-Path $summaryPath) -and (Test-Path $smtpPath)) {
    $summary = Get-Content $summaryPath -Raw | ConvertFrom-Json
    $smtpConf = Get-Content $smtpPath | ConvertFrom-Json
    $cred = New-Object System.Management.Automation.PSCredential(
      $smtpConf.user,
      (ConvertTo-SecureString $smtpConf.pass -AsPlainText -Force)
    )
    $body = $summary | ConvertTo-Json -Depth 6
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
    "Mail sent to shouichi.ikeda@gmail.com" | Out-File -FilePath $logPath -Append -Encoding utf8
  } else {
    "No smtp.json or summary not found; skip mail." | Out-File -FilePath $logPath -Append -Encoding utf8
  }
} catch {
  ("Mail send failed: {0}" -f $_.Exception.Message) | Out-File -FilePath $logPath -Append -Encoding utf8
}

("=== DailyReplay end {0} ===" -f (Get-Date -Format s)) | Out-File -FilePath $logPath -Append -Encoding utf8
