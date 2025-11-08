import win32com.client, time
WB = r"C:/AI/asagake/SHINSOKU.xlsm"
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.EnableEvents = False
excel.AutomationSecurity = 1
wb = excel.Workbooks.Open(WB, UpdateLinks=False, ReadOnly=False, AddToMru=False)
try:
    excel.Run("SHINSOKU.xlsm!AutoTrader.ResetDashboardHeaders")
    excel.Run("SHINSOKU.xlsm!AutoTrader.ButtonLoadCandidates")
    excel.Run("SHINSOKU.xlsm!AutoTrader.ButtonPushCandidates")
finally:
    wb.Close(SaveChanges=True)
    excel.Quit()
