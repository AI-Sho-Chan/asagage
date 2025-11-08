from pathlib import Path
import time
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
    excel.AutomationSecurity = 1  # allow macros
except Exception:
    pass
wb = excel.Workbooks.Open(str(wb_path))
try:
    excel.Run("AutoTrader.InstallRealtimeFormulas")
    time.sleep(0.5)
    ws = wb.Worksheets("NewDashboard")
    for addr in ["H6","I6","J6","K6","L6","M6","N6"]:
        c = ws.Range(addr)
        try:
            has_formula = bool(c.HasFormula)
        except Exception:
            has_formula = False
        print(addr, "HasFormula=", has_formula, "Formula=", c.Formula)
finally:
    wb.Close(SaveChanges=True)
    excel.Quit()

