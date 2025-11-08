import datetime as dt
from pathlib import Path

import win32com.client  # type: ignore

WB_PATH = Path(r"C:\AI\asagake\SHINSOKU.xlsm")
SUMMARY_ROWS = [
    ("Nightly Batch (16:30)", "Top150 universe, coarse → refine (ASHA + Bayes), H1/H2/H3有効化、候補CSV生成・reports更新、6h以内完走目標"),
    ("Morning Batch (05:30)", "最新1分足の取得補完、TopNユニバース更新、Excel/板/取引ログ輸出、欠損チェック"),
    ("Weekend Batch", "週末フルスキャン: parquet整合性、全セッション再評価、マスク/Optuna priors更新、バックアップ整理"),
    ("H1 (Market)", "日経騰落率に応じてJ_thを±0.10調整"),
    ("H2 (Gap)", "ギャップ帯域でJ_th加算・逆方向スキップ: 50-80:+0.1, 80-120:+0.1, 120-200:+0.2, 200-400:+0.3, 400-500:+0.4, 500-700:+0.5"),
    ("H3 (Dynamic TP/SL)", "|J|-J_thの超過分に応じてTP+0.15/SL+0.10 × max(0,|J|-|J_th|)")
]

def backup_workbook() -> Path:
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = WB_PATH.with_name(f"SHINSOKU_backup_{ts}.xlsm")
    backup.write_bytes(WB_PATH.read_bytes())
    return backup


def write_overview(ws) -> None:
    ws.Range("A1").Value = "システム概要"
    ws.Range("A2").Value = f"最終更新: {dt.datetime.now():%Y-%m-%d %H:%M:%S}"
    row = 4
    for title, desc in SUMMARY_ROWS:
        ws.Cells(row, 1).Value = title
        ws.Cells(row, 2).Value = desc
        row += 1


def main() -> None:
    backup_workbook()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(WB_PATH))
        ws = None
        for name in ("SystemOverview", "System Overview", "システム概要", "概要"):
            try:
                ws = wb.Worksheets(name)
                break
            except Exception:
                continue
        if ws is None:
            ws = wb.Worksheets.Add()
            ws.Name = "SystemOverview"
        write_overview(ws)
        wb.Close(SaveChanges=True)
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
