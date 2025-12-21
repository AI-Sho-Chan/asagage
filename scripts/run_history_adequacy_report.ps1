Param(
    [string]$Candidates = "C:/AI/asagake/output/excel/candidates_nextday.csv",
    [int]$Lookback = 90,
    [string]$Recipient = "shouichi.ikeda@gmail.com"
)

$ErrorActionPreference = 'Stop'

$repo = "C:/AI/asagake"
$python = "C:/Python313/python.exe"
$ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$logDir = Join-Path $repo "logs"
$logPath = Join-Path $logDir "history_report.log"

if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

$cmd = @(
    "tools/report_history_adequacy.py",
    "--candidates", $Candidates,
    "--lookback", $Lookback
)

"[$ts] start history adequacy report" | Out-File -FilePath $logPath -Append -Encoding utf8

$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = $python
$psi.Arguments = ($cmd -join " ")
$psi.WorkingDirectory = $repo
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError = $true
$psi.UseShellExecute = $false
$proc = New-Object System.Diagnostics.Process
$proc.StartInfo = $psi
$null = $proc.Start()
$stdout = $proc.StandardOutput.ReadToEnd()
$stderr = $proc.StandardError.ReadToEnd()
$proc.WaitForExit()
"[$ts] stdout`n$stdout" | Out-File -FilePath $logPath -Append -Encoding utf8
if ($stderr) { "[$ts] stderr`n$stderr" | Out-File -FilePath $logPath -Append -Encoding utf8 }

$jsonPath = Get-ChildItem -Path (Join-Path $repo "analysis") -Filter "history_adequacy_*.json" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $jsonPath) {
    "[$ts] no json output found; skip email" | Out-File -FilePath $logPath -Append -Encoding utf8
    exit 0
}

try {
    $summary = Get-Content $jsonPath.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
} catch {
    "[$ts] failed to parse json (UTF8): $_" | Out-File -FilePath $logPath -Append -Encoding utf8
    exit 0
}

$mailPath = Join-Path $jsonPath.DirectoryName ("{0}_mail.txt" -f [System.IO.Path]::GetFileNameWithoutExtension($jsonPath.Name))
if (Test-Path $mailPath) {
    # NOTE: Avoid Japanese literals in this .ps1 to reduce mojibake risk on older PowerShell.
    $bodyText = Get-Content $mailPath -Raw -Encoding UTF8
} else {
    $pctLtHalf = [math]::Round($summary.pct_lt_half_lb * 100, 2)
    $pctLtLb = [math]::Round($summary.pct_lt_lb * 100, 2)
    $pctGe2x = [math]::Round($summary.pct_ge_2x_lb * 100, 2)
    $body = @()
    $body += "ASAGAKE history adequacy report (JST) $ts"
    $body += "candidates: $Candidates"
    $body += ("lookback_days: {0}" -f $summary.lookback_days)
    $body += ("history_min/median/max: {0}/{1}/{2}" -f $summary.history_min, $summary.history_median, $summary.history_max)
    $body += ("pct_lt_lb: {0}%" -f $pctLtLb)
    $body += ("pct_lt_half_lb: {0}%" -f $pctLtHalf)
    $body += ("pct_ge_2x_lb: {0}%" -f $pctGe2x)
    $body += "messages:"
    foreach ($m in $summary.messages) { $body += "- $m" }
    $bodyText = $body -join "`r`n"
}

$smtpHost = $env:ASAGAKE_SMTP_HOST
$smtpPort = $env:ASAGAKE_SMTP_PORT
$smtpUser = $env:ASAGAKE_SMTP_USER
$smtpPass = $env:ASAGAKE_SMTP_PASS

# Fallback: state/smtp.json (host, port, user, pass)
if (-not $smtpHost -or -not $smtpPort -or -not $smtpUser -or -not $smtpPass) {
    $cfgPath = Join-Path $repo "state/smtp.json"
    if (Test-Path $cfgPath) {
        try {
            $cfg = Get-Content $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if (-not $smtpHost -and $cfg.host) { $smtpHost = $cfg.host }
            if (-not $smtpPort -and $cfg.port) { $smtpPort = $cfg.port }
            if (-not $smtpUser -and $cfg.user) { $smtpUser = $cfg.user }
            if (-not $smtpPass -and $cfg.pass) { $smtpPass = $cfg.pass }
        } catch {
            "[$ts] failed to read smtp.json: $_" | Out-File -FilePath $logPath -Append -Encoding utf8
        }
    }
}

if ($smtpHost -and $smtpPort -and $smtpUser -and $smtpPass) {
    try {
        $portInt = [int]$smtpPort
    } catch {
        "[$ts] mail skipped (invalid port '$smtpPort')" | Out-File -FilePath $logPath -Append -Encoding utf8
        exit 0
    }
    try {
        $secure = ConvertTo-SecureString $smtpPass -AsPlainText -Force
        $cred = New-Object System.Management.Automation.PSCredential ($smtpUser, $secure)
        Send-MailMessage `
            -From $smtpUser `
            -To $Recipient `
            -Subject "ASAGAKE history adequacy report ($ts JST)" `
            -Body $bodyText `
            -Encoding UTF8 `
            -SmtpServer $smtpHost `
            -Port $portInt `
            -UseSsl `
            -Credential $cred `
            -Attachments $jsonPath.FullName
        "[$ts] mail sent to $Recipient" | Out-File -FilePath $logPath -Append -Encoding utf8
    } catch {
        "[$ts] mail send failed: $_" | Out-File -FilePath $logPath -Append -Encoding utf8
    }
} else {
    "[$ts] mail skipped (SMTP env not set); see $($jsonPath.FullName)" | Out-File -FilePath $logPath -Append -Encoding utf8
}
