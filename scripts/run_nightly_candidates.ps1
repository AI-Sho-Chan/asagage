param(
    [switch]$Smoke = $false,
    [string]$PlanFocus = '',
    [switch]$SkipFridayWeekend = $false
)

$ErrorActionPreference = 'Stop'
$repo = 'C:/AI/asagake'
$python = ''
$logPath = Join-Path $repo 'logs/run_nightly_candidates.log'

$candidateVenv = Join-Path $repo ".venv\\Scripts\\python.exe"
if (Test-Path $candidateVenv) {
    $python = $candidateVenv
} else {
    $python = 'C:/Python313/python.exe'
}

$now = Get-Date
"[$($now.ToString('s'))] nightly_env whoami=$([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) python=$python repo=$repo" | Out-File -FilePath $logPath -Append -Encoding utf8

$legacyDisableMarker = Join-Path $repo "state\disable_local_weekend.txt"
if (Test-Path $legacyDisableMarker) {
    "[$($now.ToString('s'))] nightly_build_candidates ignore legacy marker $legacyDisableMarker" | Out-File -FilePath $logPath -Append -Encoding utf8
}

$isFriday = ($now.DayOfWeek -eq 'Friday')

if (-not $Smoke -and -not $SkipFridayWeekend -and $isFriday) {
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

if (-not $Smoke -and $SkipFridayWeekend -and $isFriday) {
    "[$($now.ToString('s'))] nightly_build_candidates skip weekend_seq (SkipFridayWeekend=1)" | Out-File -FilePath $logPath -Append -Encoding utf8
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
    '--excel','C:/AI/asagake/ASAGAKE.xlsm',
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

    # Always refresh candidates_nextday.csv so Excel Import stays on a single clean set,
    # even when nightly_build_candidates exits early (e.g. no degraded tickers).
    try {
        $aggScript = Join-Path $repo 'tools/aggregate_candidates_today.py'
        $aggOut = Join-Path $repo 'output/excel/candidates_nextday.csv'
        # Avoid overwriting candidates_nextday.csv with a tiny/partial set.
        # Keep previous "last_good" unless we can produce a reasonable number of rows.
        $aggPayload = & $python $aggScript --output $aggOut --fallback-min-rows 10 2>&1
        "[$([DateTime]::Now.ToString('s'))] aggregate_candidates_today exit 0" | Out-File -FilePath $logPath -Append -Encoding utf8
        if ($aggPayload) {
            "[$([DateTime]::Now.ToString('s'))] aggregate_candidates_today payload $aggPayload" | Out-File -FilePath $logPath -Append -Encoding utf8
        }
    }
    catch {
        "[$([DateTime]::Now.ToString('s'))] aggregate_candidates_today error $_" | Out-File -FilePath $logPath -Append -Encoding utf8
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
