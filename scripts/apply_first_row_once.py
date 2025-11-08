from __future__ import annotations

import win32com.client

excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.AutomationSecurity = 1

wb = excel.Workbooks.Open(r"C:/AI/asagake/SHINSOKU.xlsm")
try:
    ws = wb.Worksheets("NewDashboard")
    ws.Unprotect(Password="")
    formulas = {
        9: '=IF(H6="","",IFERROR(RssMarket(H6,"銘柄名称"),""))',
        10: '=IF(OR(N6="",O6="",AA6=0),"",((N6-O6)/AA6)/100)',
        11: '=IF(OR(AD6="",J6="",AD6=0),"",ABS(J6-AD6)/ABS(AD6)*100)',
        14: '=IF(H6="","",IFERROR(RssMarket(H6,"現在値"),""))',
        15: '=IF(H6="","",IFERROR(RssMarket(H6,"出来高加重平均"),""))',
        16: '=IF(H6="","",IFERROR(RssMarket(H6,"前日終値"),""))',
        17: '=IF(H6="","",IFERROR(RssMarket(H6,"気配値（買）"),""))',
        18: '=IF(H6="","",IFERROR(RssMarket(H6,"気配値（売）"),""))',
        19: '=IF(OR(Q6="",R6=""),"",(Q6+R6)/2)',
        20: '=IF(OR(S6="",P6=""),"",(S6-P6)/P6*10000)',
    }
    for col, formula in formulas.items():
        cell = ws.Cells(6, col)
        cell.Locked = False
        cell.Formula = formula
        cell.Locked = True
    wb.Save()
finally:
    wb.Close(True)
    excel.Quit()
