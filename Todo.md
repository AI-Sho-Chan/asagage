# Todo

## Approved (executing / done)
- [x] AM0945 j-cross run (coarse竊池efine) with logs and artifacts captured.
- [x] Operate FAST Top150 nightly (ASHA + shortened Bayes, H1/H3 enabled).
- [x] Continue verification of gap-aware and market-bias J adjustments.
- [x] Replace `output/excel/candidates_nextday.csv` with the filtered one-per-ticker shortlist.
- [ ] Update `SHINSOKU.xlsm` 窶懊す繧ｹ繝・Β讎りｦ≫・tab via COM with the latest batch descriptions.

## Pending approval
- [ ] Shift Task Scheduler nightly start to 16:30 (requires elevated permissions on host).
- [ ] Document the 05:30 morning batch (scripts involved, data freshness purpose).

## Backlog
- [ ] `analysis/daily_trades_*.csv` を横断集計するスクリプトを追加し、セッション別・銘柄別の実現PnLレーティングを作成する（DailyReplay のフィードバック用）。
- [ ] 時間帯分割（M0〜M3）比較検証の集計出力を追加（前提ルール: 1銘柄1ポジション、クールダウン5分、1日最大2回）。設計: `analysis/method_comparison_schema.md`
- [ ] Re-optimise gap bands vs J adders (PF / win rate / sample count / MaxDD by bucket).
- [ ] Finalise dynamic TP/SL coefficients (current seed TP:+0.15, SL:+0.10) then wire into VBA.
- [ ] Register and test the weekday 16:30 fast-nightly task (`scripts/register_fast_nightly_task.ps1`).
- [ ] Automate comparative reporting (plan-level & H1/H3 splits in `summary.xlsx`).
- [ ] Review market-volatility ﾎ寧 coefficients (B32窶釘34) as part of the weekly session tuning.
- [ ] Add conditional formatting for GapDecision (STOP=red, SKIP=orange, OK=none).
- [ ] Unify Queue/Place/Cancel/Exit logging across Orders/PnL/ExecMon.
- [x] Produce expected-P&L comparison (ﾂ･10M equal allocation) via `tools/simulate_expected_pnl.py`.
- [ ] Add Japanese display columns on NewDashboard (entry status / fill status).
- [ ] Complete straddle re-entry latch (block same-day re-fire after a loss).
- [ ] Weekend batch runtime optimisation (minute_cache deep-refresh 整理 + bt_opt30_forward early-stop 実装)
