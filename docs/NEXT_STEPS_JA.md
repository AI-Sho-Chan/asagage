# 次の手順（重要・恒久対策）

1. openpyxl 等で `SHINSOKU.xlsm`（特に NewDashboard）を直接書き換えない。
2. RSS 式の設定や修復は、VBA（`InstallRealtimeFormulas` / `SetColumnFormula` with FormulaLocal）または COM（pywin32）スクリプトで行う。喪失時はまず `python scripts/restore_dashboard_formulas.py`（Q列=「最良買気配値」、R列=「最良売気配値」、銘柄コードの `.T` を除去して数値化、L列は `NO_PRICE` 警告付き）を実行し、必要に応じて `python scripts/install_formulas.py` で VBA 側の復旧をかける。
3. 変更前に必ずバックアップ `SHINSOKU_backup_YYYYMMDD_HHMMSS.xlsm` を作成。異常が出たらバックアップから復旧してから再作業。
4. 編集後は Excel を開き、NewDashboard の I6 以降に `=RssMarket(...)` 等の式が存在し、Refresh で値が入ることを目視確認する。
5. 変更内容は `docs/handover_YYYYMMDD.md` に記録。

## 参考（Selected/Orders の仕様）
- Selected 列は、Push 後は全行 `1`（候補）。Start Auto 実行時に当日重複防止のため約定行のみ `0` へ変更。再度有効化する場合は `1` に戻すか候補 CSV を再ロード。
- ドライラン検証は `AutoTrader.PlaceOrderDryRun` を実行すると `Orders` に記録されることで確認可能。

## ドライラン検証コマンド（例）
```
python -c "import win32com.client as win32; xl=win32.DispatchEx('Excel.Application'); wb=xl.Workbooks.Open(r'C:\AI\asagake\SHINSOKU.xlsm'); xl.Visible=False; xl.DisplayAlerts=False; xl.Run('AutoTrader.PlaceOrderDryRun','9999.T','BUY',1234.5,100,'TEST'); wb.Save(); wb.Close(); xl.Quit()"
```
