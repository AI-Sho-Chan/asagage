$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location (Join-Path $scriptDir "..")
try {
    python analyze_trades.py @Args
    python scripts/build_j_stats.py --ledger analysis/all_trades_snapshot.csv --output state/j_stats.csv --min-count 12 --rules state/strategy_rules.ini --target-flat-quantile 0.8 --target-trend-quantile 0.9 --sigma-floor 0.05 --logs output/j_logs --log-days 30
    if ((Get-Date).DayOfWeek -eq 'Friday') {
        python scripts/research_trend_filters.py --output-dir analysis --lookback-days 90 --top-n 300 --sessions AM15,AM0930,AM0945,AM1000,AM1015,AM1030,PM1 --dir-threshold-bp 15 --min-count 12
    }
}
finally {
    Pop-Location
}
