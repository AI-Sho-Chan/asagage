from __future__ import annotations

from pathlib import Path

import win32com.client

WB_PATH = Path("C:/AI/asagake/SHINSOKU.xlsm")
START_ROW = 6
ROWS = 500


def jp(*codes: int) -> str:
    return "".join(chr(code) for code in codes)


RSS_FIELD_NAME = jp(0x9298, 0x67C4, 0x540D, 0x79F0)
HEADER_LAST_JP = jp(0x73FE, 0x5728, 0x5024)
RSS_FIELD_VWAP = jp(0x51FA, 0x6765, 0x9AD8, 0x52A0, 0x91CD, 0x5E73, 0x5747)
HEADER_PREV_CLOSE = jp(0x524D, 0x65E5, 0x7D42, 0x5024)
HEADER_PREOPEN_BID = jp(0x6C17, 0x914D, 0x5024, 0xFF08, 0x8CB7, 0xFF09)
HEADER_PREOPEN_ASK = jp(0x6C17, 0x914D, 0x5024, 0xFF08, 0x58F2, 0xFF09)
HEADER_PREOPEN_MID = jp(0x6C17, 0x914D, 0x5024, 0xFF08, 0x4E2D, 0x592E, 0xFF09)

FORMULAS = {
    8: '=IF(RC[15]="","",RC[15])',
    9: f'=IF(RC[-1]="","",IFERROR(RssMarket(RC[-1],"{RSS_FIELD_NAME}"),""))',
    10: '=IF(OR(RC[4]="",RC[5]="",RC[17]=0),"",((RC[4]-RC[5])/RC[17])/100)',
    11: '=IF(OR(RC[19]="",RC[-1]="",RC[19]=0),"",ABS(RC[-1]-RC[19])/ABS(RC[19])*100)',
    14: f'=IF(RC[-6]="","",IFERROR(RssMarket(RC[-6],"{HEADER_LAST_JP}"),""))',
    15: f'=IF(RC[-7]="","",IFERROR(RssMarket(RC[-7],"{RSS_FIELD_VWAP}"),""))',
    16: f'=IF(RC[-8]="","",IFERROR(RssMarket(RC[-8],"{HEADER_PREV_CLOSE}"),""))',
    17: f'=IF(RC[-9]="","",IFERROR(RssMarket(RC[-9],"{HEADER_PREOPEN_BID}"),""))',
    18: f'=IF(RC[-10]="","",IFERROR(RssMarket(RC[-10],"{HEADER_PREOPEN_ASK}"),""))',
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

        ws.Range(ws.Cells(START_ROW, 12), ws.Cells(START_ROW + ROWS, 12)).Value = ""
        ws.Range(ws.Cells(START_ROW, 13), ws.Cells(START_ROW + ROWS, 13)).Value = ""

        ws.Cells.Locked = False
        ws.Range(ws.Cells(START_ROW, 8), ws.Cells(START_ROW + ROWS, 20)).Locked = True
        ws.Protect(
            Password="",
            UserInterfaceOnly=True,
            AllowFormattingCells=True,
            AllowSorting=True,
            AllowFiltering=True,
        )

        wb.Save()
    finally:
        wb.Close(SaveChanges=True)
        excel.Quit()


if __name__ == "__main__":
    main()
