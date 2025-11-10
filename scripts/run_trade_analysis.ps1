$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

function Invoke-PythonWithRetry {
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments,
        [int]$MaxRetry = 2
    )
    for ($attempt = 1; $attempt -le $MaxRetry; $attempt++) {
        try {
            & python @Arguments
            return
        }
        catch {
            if ($attempt -ge $MaxRetry) {
                throw
            }
            Start-Sleep -Seconds 5
        }
    }
}

Push-Location (Join-Path $scriptDir "..")
try {
    $analyzeArgs = @("analyze_trades.py") + $Args
    Invoke-PythonWithRetry -Arguments $analyzeArgs

    $jStatsArgs = @(
        "scripts/build_j_stats.py",
        "--ledger","analysis/all_trades_snapshot.csv",
        "--output","state/j_stats.csv",
        "--min-count","12",
        "--rules","state/strategy_rules.ini",
        "--target-flat-quantile","0.8",
        "--target-trend-quantile","0.9",
        "--sigma-floor","0.05",
        "--logs","output/j_logs",
        "--log-days","30"
    )
    Invoke-PythonWithRetry -Arguments $jStatsArgs

    if ((Get-Date).DayOfWeek -eq 'Friday') {
        $trendArgs = @(
            "scripts/research_trend_filters.py",
            "--output-dir","analysis",
            "--lookback-days","90",
            "--top-n","300",
            "--sessions","AM15,AM0930,AM0945,AM1000,AM1015,AM1030,PM1",
            "--dir-threshold-bp","15",
            "--min-count","12"
        )
        Invoke-PythonWithRetry -Arguments $trendArgs
    }

    $rdArgs = @(
        "tools/update_rd_window_status.py",
        "--summary","analysis/session_mode_summary.csv",
        "--output","analysis/rd_window_status.csv",
        "--sessions","MID1030,PM1230"
    )
    Invoke-PythonWithRetry -Arguments $rdArgs
}
finally {
    Pop-Location
}
