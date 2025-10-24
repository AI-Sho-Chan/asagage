# Codex Operating Notes

## Excel/VBA Editing Workflow
- Treat `C:/AI/asagake/SHINSOKU.xlsm` as the production workbook. Never open Excel or the VBA editor for permanent changes.
- When changes are required, take a timestamped copy (e.g. `SHINSOKU_work_YYYYMMDD_HHMMSS.xlsm`). Modify XML or `.bas` files offline (zip/unzip), validate, then replace the production file and keep the previous version as backup.
- Use `python scripts/repair_realtime_formulas.py` to reapply dashboard formulas (I/J/K/N/O/P/Q/R/S/T columns) whenever formulas are lost.
- When AutoTrader.bas is updated, import via `python scripts/excel_install_macros.py C:/AI/asagake/SHINSOKU.xlsm C:/AI/asagake/AutoTrader.bas` instead of using the VBA editor.
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
