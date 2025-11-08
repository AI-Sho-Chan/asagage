from __future__ import annotations

from pathlib import Path

import win32com.client

WB_PATH = Path("C:/AI/asagake/SHINSOKU.xlsm")
START_ROW = 6
ROWS = 500

FORMULAS = {
    8: '=IF(RC[15]="","",RC[15])',
    9: '=IF(RC[-1]="","",IFERROR(RssMarket(RC[-1],"銘柄名称"),""))',
    10: '=IF(OR(RC[4]="",RC[5]="",RC[17]=0),"",((RC[4]-RC[5])/RC[17])/100)',
    11: '=IF(OR(RC[19]="",RC[-1]="",RC[19]=0),"",ABS(RC[-1]-RC[19])/ABS(RC[19])*100)',
    14: '=IF(RC[-6]="","",IFERROR(RssMarket(RC[-6],"現在値"),""))',
    15: '=IF(RC[-7]="","",IFERROR(RssMarket(RC[-7],"出来高加重平均"),""))',
    16: '=IF(RC[-8]="","",IFERROR(RssMarket(RC[-8],"前日終値"),""))',
    17: '=IF(RC[-9]="","",IFERROR(RssMarket(RC[-9],"気配値（買）"),""))',
    18: '=IF(RC[-10]="","",IFERROR(RssMarket(RC[-10],"気配値（売）"),""))',
    19: '=IF(OR(RC[-2]="",RC[-1]=""),"",(RC[-2]+RC[-1])/2)',
    20: '=IF(OR(RC[-1]="",RC[-4]=""),"",(RC[-1]-RC[-4])/RC[-4]*10000)',
}


def main() -> None:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        excel.AutomationSecurity = 1
    except Exception:
        pass

    wb = excel.Workbooks.Open(str(WB_PATH))
    try:
        ws = wb.Worksheets("NewDashboard")
        try:
            ws.Unprotect(Password="")
        except Exception:
            pass

        for col, formula in FORMULAS.items():
            rng = ws.Range(ws.Cells(START_ROW, col), ws.Cells(START_ROW + ROWS, col))
            rng.FormulaR1C1 = formula
            ws.Cells(START_ROW, col).FormulaR1C1 = formula
            print(col, ws.Cells(START_ROW, col).Formula)

        wb.Save()
    finally:
        wb.Close(SaveChanges=True)
        excel.Quit()


if __name__ == "__main__":
    main()
