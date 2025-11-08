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
    vbproj = wb.VBProject
    names = [vbproj.VBComponents.Item(i+1).Name for i in range(vbproj.VBComponents.Count)]
    print("modules", names)
    mod = vbproj.VBComponents("AutoTrader")
    code = mod.CodeModule.Lines(1, mod.CodeModule.CountOfLines)
    print(code[:200])
finally:
    if wb is not None:
        wb.Close(SaveChanges=False)
    excel.Quit()
