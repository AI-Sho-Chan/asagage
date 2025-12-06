param(
    [switch]$Smoke = $false,
    [string]$PlanFocus = ''
)

$ErrorActionPreference = 'Stop'
$repo = 'C:/AI/asagake'
$python = 'C:/Python313/python.exe'
$logPath = Join-Path $repo 'logs/run_nightly_candidates.log'

$DisableLocalWeekend = (Test-Path (Join-Path $repo "state\disable_local_weekend.txt"))

$now = Get-Date
$isFriday = ($now.DayOfWeek -eq 'Friday')

if (-not $Smoke -and -not $DisableLocalWeekend -and $isFriday) {
    "[$($now.ToString('s'))] nightly_build_candidates start weekend_seq" | Out-File -FilePath $logPath -Append -Encoding utf8
    try {
        & "$repo/scripts/run_weekend_then_nightly.ps1"
        $seqCode = $LASTEXITCODE
        "[$([DateTime]::Now.ToString('s'))] weekend_seq exit $seqCode" | Out-File -FilePath $logPath -Append -Encoding utf8
        if ($seqCode -ne 0) { exit $seqCode }
        try {
            & "$repo/scripts/run_trade_analysis.ps1" | Out-Null
            "[$([DateTime]::Now.ToString('s'))] run_trade_analysis completed" | Out-File -FilePath $logPath -Append -Encoding utf8
        }
        catch {
            "[$([DateTime]::Now.ToString('s'))] run_trade_analysis error $_" | Out-File -FilePath $logPath -Append -Encoding utf8
        }
        exit 0
    }
    catch {
        "[$([DateTime]::Now.ToString('s'))] weekend_seq error $_" | Out-File -FilePath $logPath -Append -Encoding utf8
        exit 1
    }
}

if (-not $Smoke -and $DisableLocalWeekend -and $isFriday) {
    "[$($now.ToString('s'))] nightly_build_candidates skip weekend_seq (disabled locally)" | Out-File -FilePath $logPath -Append -Encoding utf8
    "[$($now.ToString('s'))] nightly_build_candidates skip nightly (disabled locally)" | Out-File -FilePath $logPath -Append -Encoding utf8
    exit 0
}

$checkArgs = @('scripts/check_trading_day.py', '--date', $now.ToString('yyyy-MM-dd'))
$checkProcess = Start-Process -FilePath $python -ArgumentList $checkArgs -WorkingDirectory $repo -NoNewWindow -PassThru -Wait
if ($checkProcess.ExitCode -ne 0) {
    "[$($now.ToString('s'))] nightly_build_candidates skip (holiday)" | Out-File -FilePath $logPath -Append -Encoding utf8
    exit 0
}

$arguments = @(
    'scripts/nightly_build_candidates.py',
    '--excel','C:/AI/asagake/SHINSOKU.xlsm',
    '--base-out','output/bt30',
    '--run-type','weekday',
    '--target-date',$now.ToString('yyyyMMdd'),
    '--enable-asha',
    '--mask-ineffective','--mask-window','20','--mask-threshold','1.05',
    '--enable-bayes','--bayes-trials','16','--bayes-timeout','600',
    '--slipbp','4','--feebp','4',
    '--liquidity-quantile','0.5',
    '--jobs','4','--min-forward-ci','0.65',
    '--universe-mode','excel',
    '--enable-market-features',
    '--analysis-ledger',

    '--reopt-degraded-only',
    '--reopt-pf-th','1.2',
    '--reopt-ci-th','0.6',
    '--excel-summary'
)
if ($PlanFocus) {
    $arguments += @('--plan-focus', $PlanFocus)
}

$ts = $now.ToString('yyyyMMdd_HHmmss')
$errLog = Join-Path $repo ("logs/nightly_py_error_{0}.log" -f $ts)

$start = Get-Date
"[$($start.ToString('s'))] nightly_build_candidates start" | Out-File -FilePath $logPath -Append -Encoding utf8
try {
    $process = Start-Process -FilePath $python `
        -ArgumentList $arguments `
        -WorkingDirectory $repo `
        -NoNewWindow `
        -Wait `
        -PassThru `
        -RedirectStandardError $errLog
    $code = $process.ExitCode
    $end = Get-Date
    "[$($end.ToString('s'))] nightly_build_candidates exit $code" | Out-File -FilePath $logPath -Append -Encoding utf8
    if ($code -ne 0 -and (Test-Path $errLog)) {
        "[$($end.ToString('s'))] nightly_build_candidates error_detail $errLog" | Out-File -FilePath $logPath -Append -Encoding utf8
    }
    if (-not $Smoke -and $code -eq 0) {
        try {
            & "$repo/scripts/run_trade_analysis.ps1" | Out-Null
            "[$($end.ToString('s'))] run_trade_analysis completed" | Out-File -FilePath $logPath -Append -Encoding utf8
        }
        catch {
            "[$($end.ToString('s'))] run_trade_analysis error $_" | Out-File -FilePath $logPath -Append -Encoding utf8
        }
    }
    exit $code
}
catch {
    $end = Get-Date
    "[$($end.ToString('s'))] nightly_build_candidates error $_" | Out-File -FilePath $logPath -Append -Encoding utf8
    throw
}


