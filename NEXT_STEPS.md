# Project Status

- Archived on `2026-03-25`.
- Final revalidation result: the actual-like replay on `2026-03-20` to `2026-03-25` was `-847,331 yen` overall and `-373,499 yen` for `LIVE_STRONG`.
- The current simple replay looked better than the actual-like replay, so the previous validation path was too optimistic.
- Do not resume development or trading operations without explicit re-approval and a new hypothesis.
- Read first:
  - `docs/project_archive_20260325.md`
  - `reports/revalidation_report_march20260320_0325.md`
  - `reports/revalidation_report_feb20260209_0218.md`

## Historical Checklist

1. Confirm the restored workbook opens without the repair prompt. If Excel offers recovery, cancel it and verify the ribbon buttons are present.
2. Run `python scripts/repair_realtime_formulas.py` once to repopulate NewDashboard formulas after confirming the workbook state.
3. Re-open `AutoTrader.bas` (export) on a working copy and re-apply the planned improvements (SetColumnFormula refactor, EnsureRealtimeFirstRow logic, button automation) following the offline editing rules in `docs/codex.md`.
4. After modifications, use `scripts/swap_workbook.py` to replace production safely and log the change in `docs/handover_20251020.md`.
5. Re-test buttons (`Load Candidates`, `Push Candidates`, `Refresh Now`, `Start Auto`) and capture any failures with `logs/autotrader_debug.log`.
6. Once validated, remove temporary files such as `SHINSOKU_corrupted_*.xlsm` if no longer needed.

---

## 2025-10-26 Automation & Data Pipeline Update

- 1分足データ統合  
  - `scripts/bt_opt30_forward.py` に `--use-local-raw` を実装し、ナイトバッチ (`scripts/nightly_build_candidates.py`) では常時ON。ローカル保存済みの 1 分足を主、Yahoo補完を従にする運用へ移行。  
  - 全銘柄対象の日次保存ジョブ `scripts/update_all_1m.py` を強化（30日制限の扱い・欠損ログ追加）し、毎平日 05:30 に自動実行（`scripts/register_update_1m_task.ps1`）。
- 板ログ／約定ログの自動取得  
  - `excel/BoardLogger.xlsx` をテンプレ化し、10本板セルに楽天RSSを設定。  
  - スナップショットを毎平日 08:55 起動のデーモン (`scripts/board_logger_daemon.py` + `scripts/run_board_logger_daemon.cmd`) で 09:00–15:30 の間1分刻みで `output/board_logs/YYYYMMDD/` に保存。  
  - 取引ログ (`Orders`, `PnL`, `ExecMon`, `NewDashboard`) はナイト後タスクで毎朝07:20自動エクスポート。
- ナイト後ポスト処理とサイズ計画  
  - `scripts/post_nightly_tasks.py` が `export_excel_logs.py` と `make_size_plan.py` を連動。  
  - タスク `Asagake-PostNightly`（07:20）と `Asagake-BoardLogger`（08:55/平日）、`Asagake-Update1m`（05:30/平日）をPowerShellスクリプト経由で登録済み。
- サイズ最適化＆比較検証  
  - `scripts/portfolio_sizing_sim.py` で動的サイズとベースラインの比較を実施、結果は `output/research/portfolio_sim/<timestamp>/`。  
  - ローカルvsリモート再取得比較用 `scripts/run_local_vs_remote_compare.py` を追加し、直近分は `output/_COMPARE_SUMMARY.csv` に集約。
- すべての出力は既存成果物に影響を与えず、サイドディレクトリ（`output/trade_logs/`, `output/board_logs/`, `output/excel/size_plan/` 等）に隔離。既存Excel/VBAは無変更。
