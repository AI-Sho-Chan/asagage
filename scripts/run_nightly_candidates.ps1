param()

$ErrorActionPreference = 'Stop'
$repo = 'C:\AI\asagake'
$python = 'C:\Python313\python.exe'
$logPath = Join-Path $repo 'logs\run_nightly_candidates.log'

$now = Get-Date
if ($now.DayOfWeek -eq 'Friday') {
    "[$($now.ToString('s'))] nightly_build_candidates skip (Friday run disabled)" | Out-File -FilePath $logPath -Append -Encoding utf8
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
    '--excel-summary'
)

$start = Get-Date
"[$($start.ToString('s'))] nightly_build_candidates start" | Out-File -FilePath $logPath -Append -Encoding utf8
try {
    $process = Start-Process -FilePath $python -ArgumentList $arguments -WorkingDirectory $repo -NoNewWindow -Wait -PassThru
    $code = $process.ExitCode
    $end = Get-Date
    "[$($end.ToString('s'))] nightly_build_candidates exit $code" | Out-File -FilePath $logPath -Append -Encoding utf8
    if ($code -eq 0) {
        try {
            & "$repo\scripts\run_trade_analysis.ps1" | Out-Null
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
