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

$args = @(
  "tools/simulate_daily_replay.py",
  "--date", $DateTag,
  "--nominal", [string][int][math]::Round($Nominal)
)

& $python $args

# メール送信（state/smtp.json がある場合のみ）
try {
  $summaryPath = Join-Path $repo ("analysis/daily_replay_{0}.json" -f $DateTag)
  $smtpPath = Join-Path $repo "state/smtp.json"
  if (Test-Path $summaryPath -and Test-Path $smtpPath) {
    $summary = Get-Content $summaryPath -Raw | ConvertFrom-Json
    $smtpConf = Get-Content $smtpPath | ConvertFrom-Json
    $cred = New-Object System.Management.Automation.PSCredential(
      $smtpConf.user,
      (ConvertTo-SecureString $smtpConf.pass -AsPlainText -Force)
    )
    $body = $summary | ConvertTo-Json -Depth 3
    Send-MailMessage `
      -To "shouichi.ikeda@gmail.com" `
      -From $smtpConf.user `
      -Subject ("ASAGAKE DailyReplay {0}" -f $DateTag) `
      -Body $body `
      -SmtpServer $smtpConf.host `
      -Port [int]$smtpConf.port `
      -UseSsl `
      -Credential $cred
    Write-Host "Mail sent to shouichi.ikeda@gmail.com"
  } else {
    Write-Host "No smtp.json or summary not found; skip mail."
  }
} catch {
  Write-Host ("Mail send failed: {0}" -f $_.Exception.Message)
}
