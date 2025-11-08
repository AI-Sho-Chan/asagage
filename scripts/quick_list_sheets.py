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
try:
    wb = excel.Workbooks.Open(str(wb_path))
except Exception as e:
    print("OPEN_ERROR", e)
    excel.Quit()
    sys.exit(2)

print("Workbook:", wb_path)
for i in range(1, wb.Worksheets.Count + 1):
    try:
        print("SHEET", i, wb.Worksheets(i).Name)
    except Exception:
        pass

wb.Close(SaveChanges=False)
excel.Quit()
