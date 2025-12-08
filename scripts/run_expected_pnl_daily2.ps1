param(
  [double]$Budget = 10000000
)

$ErrorActionPreference = "Stop"
$repo = "C:/AI/asagake"
$python = "C:/Python313/python.exe"

Set-Location $repo

$today   = Get-Date
$dateTag = $today.ToString("yyyyMMdd")

$candPath = Join-Path $repo "output/excel/candidates_nextday.csv"
if (-not (Test-Path $candPath)) {
    Write-Host "candidates_nextday.csv not found; skipping expected PnL simulation."
    exit 0
}

$outDir = Join-Path $repo "analysis"
if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}

$jsonPath = Join-Path $outDir ("expected_pnl_{0}.json" -f $dateTag)
$csvPath  = Join-Path $outDir "expected_pnl_daily.csv"

$args = @(
    "tools/simulate_expected_pnl.py",
    "--budget", [string][int][math]::Round($Budget),
    "--current", "output/excel/candidates_nextday.csv"
)

try {
    $jsonText = & $python $args
} catch {
    Write-Host ("simulate_expected_pnl.py failed: {0}" -f $_.Exception.Message)
    exit 1
}

if ([string]::IsNullOrWhiteSpace($jsonText)) {
    Write-Host "simulate_expected_pnl.py produced no output; skipping."
    exit 0
}

# Save JSON snapshot
$jsonText | Out-File -FilePath $jsonPath -Encoding utf8

try {
    $obj = $jsonText | ConvertFrom-Json
} catch {
    Write-Host "Failed to parse JSON from simulate_expected_pnl.py"
    exit 1
}

$record = [PSCustomObject]@{
    Date        = $today.ToString("yyyy-MM-dd")
    Budget      = [double]$obj.budget
    Positions   = [int]$obj.current.positions
    ExpectedYen = [double]$obj.current.expected_yen
    PerPosition = [double]$obj.current.per_position
}

if (Test-Path $csvPath) {
    $record | Export-Csv -Path $csvPath -Append -NoTypeInformation -Encoding UTF8
} else {
    $record | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
}

Write-Host "Expected PnL simulation written to:"
Write-Host ("  {0}" -f $jsonPath)
Write-Host ("  {0}" -f $csvPath)

# --- Optional email notification ---
try {
    $smtpPath = Join-Path $repo "state/smtp.json"
    if (Test-Path $smtpPath) {
        $smtpConf = Get-Content $smtpPath | ConvertFrom-Json
        $cred = New-Object System.Management.Automation.PSCredential(
            $smtpConf.user,
            (ConvertTo-SecureString $smtpConf.pass -AsPlainText -Force)
        )
        $body = @"
Date:        $($record.Date)
Budget:      $([math]::Round($record.Budget,0)) JPY
Positions:   $($record.Positions)
ExpectedPnL: $([math]::Round($record.ExpectedYen,2)) JPY
PerPosition: $([math]::Round($record.PerPosition,2)) JPY

Source: tools/simulate_expected_pnl.py (run_expected_pnl_daily2.ps1)
"@
        Send-MailMessage `
            -To "shouichi.ikeda@gmail.com" `
            -From $smtpConf.user `
            -Subject ("ASAGAKE ExpectedPnL {0}" -f $dateTag) `
            -Body $body `
            -SmtpServer $smtpConf.host `
            -Port [int]$smtpConf.port `
            -UseSsl `
            -Credential $cred
        Write-Host "Notification email sent to shouichi.ikeda@gmail.com"
    } else {
        Write-Host "smtp.json not found; skipping email notification."
    }
} catch {
    Write-Host ("Email notification failed: {0}" -f $_.Exception.Message)
}
