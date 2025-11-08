# Agent Operating Guide (Repo-wide)

This file applies to the entire repository. Codex must read and follow this on every run.

Core rules for Excel workbook (SHINSOKU.xlsm):

- Do not modify `SHINSOKU.xlsm` with libraries that cannot preserve Excel add-in formulas or protection (e.g., openpyxl). Never write formulas or values into `NewDashboard` with openpyxl or similar. This previously erased `=RssMarket(...)` formulas in `I6` onward.
- Only use Excel/VBA or COM automation for formula work:
  - VBA: `AutoTrader.InstallRealtimeFormulas`, `SetColumnFormula` (with `FormulaLocal` fallback).
  - COM (pywin32): scripts such as `scripts/repair_realtime_formulas.py` or `scripts/burn_realtime_formulas.py`.
- Always create a timestamped backup (`SHINSOKU_backup_YYYYMMDD_HHMMSS.xlsm`) before any workbook change. Restore from backup if RSS cells disappear.
- After changes, open Excel and verify `NewDashboard` columns I–T from row 6 contain live formulas and produce values when refreshed.

Must-read docs each session:

- `NEXT_STEPS.md`
- `docs/codex.md`
- Latest handover log(s), e.g., `docs/handover_20251027.md`

Operational notes:

- Selected column policy: After “Push Candidates”, all rows are `1` (candidates). “Start Auto” sets executed rows to `0` to prevent duplicate same-day orders. Re-enable by setting `1` or reloading the candidate CSV.
- Orders sheet logging: `AutoTrader.PlaceOrderDryRun` writes to `Orders` for verification. Keep this as a smoke test when validating signals.

Guardrails for contributions:

- Do not reintroduce scripts that use openpyxl to write to `.xlsm`, especially `NewDashboard` ranges. A pre-commit guard script is provided in `scripts/guard_no_openpyxl_xlsm.py`.
- If you need to touch formulas, prefer the established VBA/COM paths and log steps in a new `docs/handover_YYYYMMDD.md` entry.
- Use `scripts/restore_dashboard_formulas.py` when the I-V columns lose their formulas. This script reapplies the canonical `RssMarket` expressions over 600 rows (Q="最良買気配値", R="最良売気配値"、Ticker の `.T` サフィックスを除去して数値化) and re-protects the sheet.
- L/M 列はシート側で `Selected=1` かつ `|J| >= |J_th|` のときに BUY/SELL とモード（例: `BUY / j-only`）を表示し、必要な価格がない場合は `NO_PRICE` を警告する数式を設定済み。VBA が別途書き込む際はその値が優先される。
