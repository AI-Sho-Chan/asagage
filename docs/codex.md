# Codex Operating Notes

## Excel/VBA Editing Workflow
- Treat `C:/AI/asagake/SHINSOKU.xlsm` as the production workbook. Never open Excel or the VBA editor for permanent changes.
- Confirm `ASAGAKE.xlsm` is fully closed (no Excel/VBE instances) before running repair/install scripts. If it might be open, pause and coordinate before proceeding.
- When changes are required, take a timestamped copy (e.g. `SHINSOKU_work_YYYYMMDD_HHMMSS.xlsm`). Modify XML or `.bas` files offline (zip/unzip), validate, then replace the production file and keep the previous version as backup.
- Use `python scripts/repair_realtime_formulas.py` to reapply dashboard formulas (I/J/K/N/O/P/Q/R/S/T columns) whenever formulas are lost.
- Ensure `logs/vba_events.log` exists and is writable so every macro error path can append diagnostics for later review.
- When AutoTrader.bas is updated, import via `python scripts/excel_install_macros.py C:/AI/asagake/SHINSOKU.xlsm C:/AI/asagake/AutoTrader.bas` instead of using the VBA editor.
- Run `python scripts/auto_repair_asagake.py --skip-formulas` after batches to reapply the dashboard layout and re-import AutoTraderAdvanced / cDashboardWatcher without opening Excel.
- For dashboard構成の修正や V2 レイアウト更新は、必ず `python scripts/repair_asagake_dashboard.py --excel C:/AI/asagake/ASAGAKE.xlsm` → `python scripts/excel_install_macros.py C:/AI/asagake/ASAGAKE.xlsm excel/AutoTraderAdvanced.bas excel/cDashboardWatcher.cls` の手順で行い、Excel/VBE は開かない（貼付エラー防止のため）。
- Excel を閉じた状態で `repair_asagake_dashboard.py` と `excel_install_macros.py` を実行し、完了後に Excel を再起動して反映を確認すること。ASAGAKE.xlsm の更新時刻と `cDashboardWatcher` が VBAProject に存在するかを確認する。
- NewDashboardV2 の AP～AT 行タイトルが文字化けした場合は、`scripts/repair_asagake_dashboard.py` の `JP_MAP` を UTF-8 で修正したうえで repair → macro install の順に再実行する。旧 `JP_MAP`（mojibake）のままだと今回のように変更が反映されない。
- Always log the source backup and the applied changes in `docs/handover_YYYYMMDD.md`.

## Backup Expectations
- Before each workbook edit, create `SHINSOKU_YYYYMMDD_HHMMSS.xlsm` in the same directory.
- After modifications, archive the previous production file (e.g. rename to `SHINSOKU_corrupted_...` when rolling back).
- Preserve `AutoTrader.bas.bak` as the canonical rollback for the VBA module.

## Recovery Scripts
- `scripts/repair_realtime_formulas.py`: reinstalls dashboard formulas and re-applies sheet protection.
- `scripts/verify_repair.py`: debug helper to print applied formulas without opening Excel.
- `scripts/swap_workbook.py`: safe swap between a patched workbook and production (used when Excel must stay closed).

## Communication
- Whenever Excel/VBA had to be avoided, note the exact steps (copy source, modify XML, replacement) in both `handover_YYYYMMDD.md` and future chat instructions.

## Automation Principles
- Tradingシステムは「取引開始」「取引終了」以外の工程（候補生成・学習・適用・可視化・アラート・ローテーション）をフル自動で回すことを前提に設計・実装する。手動オペが発生した場合は恒久的な自動化策を直ちに検討し、次回以降は人手不要にする。
- 週末/ナイトなど長時間ジョブがマーケットオープンに間に合わない兆候を検知した時点で、必ず時間短縮の代替案（停止→時短再実行、グリッド縮小、残時間見積など）を提示し、ユーザー判断を仰ぐ。

## ASAGAKE 移行メモ（重要）
- 本番ワークブックは現在 `C:/AI/asagake/ASAGAKE.xlsm` を使用します。過去ドキュメントの `SHINSOKU.xlsm` 記載はレガシーです。
- バッチ実行系（`run_weekend_then_nightly.ps1`, `run_weekly_screening.ps1`, `nightly_build_candidates.py`）は ASAGAKE 参照に統一済み。
- ダッシュボード修復系も `scripts/restore_dashboard_formulas.py` を ASAGAKE 既定に修正済み。バックアップ名も `ASAGAKE_backup_YYYYMMDD_HHMMSS.xlsm` へ統一しました。
- なお一部の保守用スクリプトは既定で `SHINSOKU.xlsm` を指す箇所が残存します（手動・個別検証用）。バッチフローでは呼ばれません。誤実行を防ぐため、運用時は `--excel C:/AI/asagake/ASAGAKE.xlsm` を明示するか、該当スクリプトを使用しないでください。

## Execution Platforms
- Github Actions�i���������i�[�j�͐��\/�Z���ԉ��̖ړI�ł͎g�p���Ȃ��B���x�E���肪�K�v�Ȏ������s�̓��[�J���������� AWS HPC�iRay+Spot �Ȃǂ̍����\�\���j��O��Ƃ���B

