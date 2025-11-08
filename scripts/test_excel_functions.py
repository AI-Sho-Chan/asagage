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
    c = ws.Range("Z6")
    c.Formula = "=LET(x,1,x+1)"
    print("Z6 Value=", c.Value)
    c2 = ws.Range("Z7")
    c2.Formula = "=SUMPRODUCT({1,2,3},{4,5,6})"
    print("Z7 Value=", c2.Value)
finally:
    wb.Close(SaveChanges=False)
    excel.Quit()

