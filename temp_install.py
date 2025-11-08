import win32com.client as win32
from pathlib import Path
path = Path(r"C:\AI\asagake\SHINSOKU.xlsm")
excel = win32.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(str(path))
try:
    excel.Run('AutoTrader.InstallRealtimeFormulas')
    wb.Save()
finally:
    wb.Close(SaveChanges=True)
    excel.Quit()
