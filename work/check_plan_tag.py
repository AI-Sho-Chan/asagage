import win32com.client
from pathlib import Path
path = Path(r"c:/AI/asagake/work/SHINSOKU_test_20251104_095612.xlsm")
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.AutomationSecurity = 3
wb = None
try:
    wb = excel.Workbooks.Open(str(path), UpdateLinks=False, ReadOnly=False, AddToMru=False)
    mod = wb.VBProject.VBComponents("AutoTrader")
    code = mod.CodeModule.Lines(1, mod.CodeModule.CountOfLines)
    print('IsPlanTagAllowed present?', 'IsPlanTagAllowed' in code)
finally:
    if wb is not None:
        wb.Close(SaveChanges=False)
    excel.Quit()
