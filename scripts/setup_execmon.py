from __future__ import annotations

from pathlib import Path
import sys

try:
    import win32com.client  # type: ignore
except Exception as e:
    print("PYWIN32_IMPORT_ERROR", e)
    sys.exit(1)

WB_PATH = Path("C:/AI/asagake/SHINSOKU.xlsm")
SHEET_NAME = "ExecMon"

# RssExecutionList expects a header row that specifies the fields to output.
# We choose a minimal, generally available set of item names (Japanese):
#  約定日, 銘柄コード, 売買, 約定数量, 約定価格, 注文ID
HEADERS = ("約定日", "銘柄コード", "売買", "信用区分", "約定数量", "約定価格", "注文ID")


def main():
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        excel.AutomationSecurity = 3  # disable auto macros while wiring
    except Exception:
        pass
    wb = excel.Workbooks.Open(str(WB_PATH))
    try:
        try:
            ws = wb.Worksheets(SHEET_NAME)
        except Exception:
            ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            ws.Name = SHEET_NAME

        # Write headers A1.. in order
        for i, h in enumerate(HEADERS, start=1):
            ws.Cells(1, i).Value = h

        # RssExecutionList(header_range, 注文種類, 銘柄コード, 口座区分, 信用区分, 売買)
        # We leave filters empty to fetch all for the day; Rakuten RSS will populate under the headers.
        header_ref = ws.Range(ws.Cells(1, 1), ws.Cells(1, len(HEADERS)))
        ws.Cells(2, 1).Formula = f"=RssExecutionList({header_ref.Address})"

        wb.Save()
        print("ExecMon sheet prepared with RssExecutionList trigger.")
    finally:
        wb.Close(SaveChanges=True)
        excel.Quit()


if __name__ == "__main__":
    main()
