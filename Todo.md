# Todo

## Approved (executing / done)
- [x] AM0945 j-cross run (coarse→refine) with logs and artifacts captured.
- [x] Operate FAST Top150 nightly (ASHA + shortened Bayes, H1/H3 enabled).
- [x] Continue verification of gap-aware and market-bias J adjustments.
- [x] Replace `output/excel/candidates_nextday.csv` with the filtered one-per-ticker shortlist.
- [ ] Update `SHINSOKU.xlsm` “システム概要” tab via COM with the latest batch descriptions.

## Pending approval
- [ ] Shift Task Scheduler nightly start to 16:30 (requires elevated permissions on host).
- [ ] Document the 05:30 morning batch (scripts involved, data freshness purpose).

## Backlog
- [ ] Re-optimise gap bands vs J adders (PF / win rate / sample count / MaxDD by bucket).
- [ ] Finalise dynamic TP/SL coefficients (current seed TP:+0.15, SL:+0.10) then wire into VBA.
- [ ] Register and test the weekday 16:30 fast-nightly task (`scripts/register_fast_nightly_task.ps1`).
- [ ] Automate comparative reporting (plan-level & H1/H3 splits in `summary.xlsx`).
- [ ] Review market-volatility ΔJ coefficients (B32–B34) as part of the weekly session tuning.
- [ ] Add conditional formatting for GapDecision (STOP=red, SKIP=orange, OK=none).
- [ ] Unify Queue/Place/Cancel/Exit logging across Orders/PnL/ExecMon.
- [x] Produce expected-P&L comparison (¥10M equal allocation) via `tools/simulate_expected_pnl.py`.
- [ ] Add Japanese display columns on NewDashboard (entry status / fill status).
- [ ] Complete straddle re-entry latch (block same-day re-fire after a loss).
