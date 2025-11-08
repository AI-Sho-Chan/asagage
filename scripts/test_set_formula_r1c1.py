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
excel.AutomationSecurity = 1
wb = excel.Workbooks.Open(str(wb_path))
ws = wb.Worksheets("NewDashboard")
try:
    c = ws.Range("I6")
    try:
        c.FormulaR1C1 = '=IF(RC[-1]="","",IFERROR(RssMarket(RC[-1],3),""))'
        ok = True
    except Exception as e:
        print("SET_R1C1_ERROR", e)
        ok = False
    print("I6 HasFormula=", bool(c.HasFormula), "Formula=", c.Formula)
finally:
    wb.Close(SaveChanges=True)
    excel.Quit()

