import time
import win32com.client
WB_PATH = r"C:/AI/asagake/SHINSOKU.xlsm"
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.EnableEvents = False
try:
    try:
        excel.AutomationSecurity = 1
    except Exception:
        pass
    wb = excel.Workbooks.Open(WB_PATH)
    try:
        excel.Run("AutoTrader.ResetDashboardHeaders")
        time.sleep(0.5)
        excel.Run("AutoTrader.ButtonLoadCandidates")
        time.sleep(0.5)
        excel.Run("AutoTrader.ButtonPushCandidates")
        time.sleep(0.5)
        excel.Run("AutoTrader.ButtonRefreshNow")
        time.sleep(0.5)
    finally:
        wb.Close(SaveChanges=True)
finally:
    excel.Quit()
