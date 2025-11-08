from pathlib import Path
import sys

try:
    import win32com.client  # type: ignore
except Exception as e:
    print("PYWIN32_IMPORT_ERROR", e)
    sys.exit(1)

wb_path = Path("C:/AI/asagake/SHINSOKU.xlsm")
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(str(wb_path))
ws = wb.Worksheets("NewDashboard")
try:
    hdr_row = 5
    start_col = 1
    end_col = ws.Cells(hdr_row, ws.Columns.Count).End(1).Column  # xlToLeft=1
    for c in range(start_col, end_col + 1):
        v = ws.Cells(hdr_row, c).Value
        print(c, v if v is not None else "")
finally:
    wb.Close(SaveChanges=False)
    excel.Quit()

