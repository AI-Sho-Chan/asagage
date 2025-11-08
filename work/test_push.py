import win32com.client
from pathlib import Path
path = Path(r"c:/AI/asagake/work/SHINSOKU_test_20251104_095612.xlsm")
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.AutomationSecurity = 1
wb = None
try:
    wb = excel.Workbooks.Open(str(path), UpdateLinks=False, ReadOnly=False, AddToMru=False)
    excel.Run("AutoTrader.ButtonLoadCandidates")
    try:
        excel.Run("AutoTrader.ButtonPushCandidates")
        print("Push succeeded")
    except Exception as exc:
        print("Push failed", exc)
        raise
finally:
    if wb is not None:
        wb.Close(SaveChanges=False)
    excel.Quit()
