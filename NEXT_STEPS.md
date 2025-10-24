# Next Steps Checklist

1. Confirm the restored workbook opens without the repair prompt. If Excel offers recovery, cancel it and verify the ribbon buttons are present.
2. Run `python scripts/repair_realtime_formulas.py` once to repopulate NewDashboard formulas after confirming the workbook state.
3. Re-open `AutoTrader.bas` (export) on a working copy and re-apply the planned improvements (SetColumnFormula refactor, EnsureRealtimeFirstRow logic, button automation) following the offline editing rules in `docs/codex.md`.
4. After modifications, use `scripts/swap_workbook.py` to replace production safely and log the change in `docs/handover_20251020.md`.
5. Re-test buttons (`Load Candidates`, `Push Candidates`, `Refresh Now`, `Start Auto`) and capture any failures with `logs/autotrader_debug.log`.
6. Once validated, remove temporary files such as `SHINSOKU_corrupted_*.xlsm` if no longer needed.
