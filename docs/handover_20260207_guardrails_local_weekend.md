# Handover 2026-02-07 (Guardrails + Local Weekend)

## Scope
- Implemented operational guardrails requested for practical validation:
  - candidate freshness guard
  - daily entry rotation cap guard
- Executed weekend batch locally and imported fresh candidates to ASAGAKE.
- Switched Friday nightly flow to local-first behavior (cloud-disable marker is now ignored).

## Workbook Safety / Backup
- Before VBA import, backup was created:
  - `ASAGAKE_backup_20260207_152004.xlsm`
- Updated module in workbook:
  - `excel/AutoTraderAdvanced.bas` imported into `ASAGAKE.xlsm` via `scripts/update_asagake_vba.ps1`.

## VBA Changes
- File: `excel/AutoTraderAdvanced.bas`
- Added candidate freshness controls:
  - `CANDIDATE_MAX_AGE_HOURS_DEFAULT = 72`
  - `ResolveCandidatesCsvPathV2`
  - `IsCandidateFeedFreshV2`
- Added daily cap controls:
  - `DAILY_ENTRY_CAP_DEFAULT = 20`
  - `CountTodayEntryRowsV2`
  - `IsDailyEntryCapReachedV2`
- Applied guards:
  - `StartDemoV2` / `StartLiveV2` block start when candidates are stale.
  - `PreplaceOrdersV2` blocks new preplace and marks rows as:
    - `BLOCKED_STALE_CANDIDATE`
    - `BLOCKED_DAILY_CAP`
- Import metadata tracking added at end of `ImportCandidatesV2`.

## Rule Config Changes
- File: `state/strategy_rules.ini`
- Added:
  - `candidate_max_age_hours=72`
  - `daily_entry_cap=20`

## Weekend Batch (Local) Result
- Local weekend incremental run completed successfully.
- Status source: `logs/nightly_status.txt`
- Key values:
  - `run_type=weekend`
  - `target_date=20260206`
  - `state=success`
  - `completed_plans=12`
  - `total_candidates=27`
  - `unique_tickers=19`
  - `weekend_reopt_codes=6`
  - `weekend_keep_codes=94`
  - `weekend_reused_rows=23`
  - `elapsed_seconds=5785` (~1h36m)
- Candidates import confirmation:
  - `logs/vba_events.log` contains `ImportCandidatesV2 ... imported=27`.

## Friday Local Operation Changes
- File: `scripts/run_nightly_candidates.ps1`
  - Added `-SkipFridayWeekend` switch (default: local Friday weekend sequence runs).
  - Legacy marker `state/disable_local_weekend.txt` is now logged and ignored.
- File: `scripts/run_weekend_then_nightly.ps1`
  - Added `-WeekendMonthlyReset` switch (default OFF).
  - `--weekend-monthly-reset` is now opt-in, not always-on.
- Syntax checked:
  - `scripts/run_nightly_candidates.ps1 : OK`
  - `scripts/run_weekend_then_nightly.ps1 : OK`

## Notes
- This change is intended to reduce runtime variance and avoid accidental full resets.
- If a monthly full reset is needed, run:
  - `scripts/run_weekend_then_nightly.ps1 -WeekendMonthlyReset`

## Hotfix (VBA 1004 in UpdateStatusV2)
- User observed runtime error:
  - `Err 1004: Range クラスの HorizontalAlignment プロパティを設定できません。`
- Root context:
  - old `UpdateStatusV2` formatting path could hard-fail depending on sheet state.
- Fixes applied:
  - `UpdateStatusV2` now:
    - unprotects/reprotects safely,
    - wraps status-cell formatting in fail-safe path,
    - logs and continues on formatting issues.
  - `StartDemoV2/StopDemoV2/StartLiveV2/StopLiveV2` now have explicit error handlers with `LogVbaEvent`.
- Applied to workbook with backup via:
  - `scripts/update_asagake_vba.ps1 -Force`
  - backup: `ASAGAKE_backup_20260207_202147.xlsm`

## Hotfix (Import error 13 in PreplaceOrdersV2)
- User observed runtime error:
  - `Err 13: 型が一致しません`
  - break location in `PreplaceOrdersV2` around `hasBuy = IsNumeric(tmpVal) And tmpVal > 0`
- Cause:
  - `tmpVal` can be Excel error variants (`#N/A` etc.), and direct comparison `tmpVal > 0` raises type mismatch.
- Fixes applied:
  - Replaced checks with safe numeric conversion:
    - `hasBuy = (ToDouble(tmpVal, 0#) > 0#)`
    - `hasSell = (ToDouble(tmpVal, 0#) > 0#)`
  - Added error handlers in `PreplaceOrdersV2`:
    - row-level logging and continue (`row_err`)
    - procedure-level logging (`Err ...`)
  - Added column-missing/error tolerance in `ApplyDynamicSignalsV2` for `Gap_bp` and other optional columns.
- Applied to workbook:
  - backup: `ASAGAKE_backup_20260207_203935.xlsm`
  - `ImportCandidatesV2` smoke test via COM: succeeded (`ok True`).

## Dynamic-J Retest Executed (bias 0.0 vs 1.0)
- Re-ran replay across 32 dates using:
  - `tools/simulate_daily_replay.py --dynamic-jth --bias-bp {0.0,1.0}`
- Outputs:
  - `analysis/daily_replay_dynamic_jth_retest_compare_20260207.csv`
  - `analysis/daily_replay_dynamic_jth_retest_summary_20260207.csv`
  - `analysis/daily_replay_dynamic_jth_retest_report_20260207.md`
- Summary:
  - `bias_bp=0.0`: `200 trades / -1,999,075 yen`
  - `bias_bp=1.0`: `199 trades / -2,096,071 yen`
  - delta (`0.0 - 1.0`): `+1 trade / +96,996 yen`
