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
for addr in ["H6","I6","J6","K6","L6","M6","N6","O6","P6","Q6","R6","S6","T6","U6","V6","AQ6","AS6","AT6","AW6"]:
    c = ws.Range(addr)
    try:
        has_formula = bool(c.HasFormula)
    except Exception:
        has_formula = False
    print(addr, "HasFormula=", has_formula, "Formula=", c.Formula, "Value=", c.Value)
wb.Close(SaveChanges=False)
excel.Quit()
