import win32com.client as win32

BOARD_RANGE_ROWS = 600
START_ROW = 6

COLUMN_FORMULAS = {
    "I": '=IF($H{row}="","",IFERROR(RssMarket($H{row},"銘柄名称"),""))',
    "J": '=IF(OR($N{row}="",$O{row}="",$AA{row}=0),"",(($N{row}-$O{row})/$AA{row})/100)',
    "K": '=IF(OR($AD{row}="",$J{row}="",$AD{row}=0),"",ABS($J{row}-$AD{row})/ABS($AD{row})*100)',
    "L": '=""',
    "M": '=""',
    "N": '=IF($H{row}="","",IFERROR(RssMarket($H{row},"現在値"),""))',
    "O": '=IF($H{row}="","",IFERROR(RssMarket($H{row},"出来高加重平均"),""))',
    "P": '=IF($H{row}="","",IFERROR(RssMarket($H{row},"前日終値"),""))',
    "Q": '=IF($H{row}="","",IFERROR(RssMarket($H{row},"気配値（買）"),""))',
    "R": '=IF($H{row}="","",IFERROR(RssMarket($H{row},"気配値（売）"),""))',
    "S": '=IF(OR(Q{row}="",R{row}=""),"",(Q{row}+R{row})/2)',
    "T": '=IF(OR(S{row}="",P{row}=""),"",(S{row}-P{row})/P{row}*10000)',
}


def set_column(ws, col_letter, formula_template):
    first = f"{col_letter}{START_ROW}"
    last = f"{col_letter}{START_ROW + BOARD_RANGE_ROWS - 1}"
    formula = formula_template.format(row=START_ROW)
    ws.Range(first).FormulaLocal = formula
    ws.Range(first).AutoFill(Destination=ws.Range(f"{first}:{last}"))


def main() -> None:
    xl = win32.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(r"C:\AI\asagake\SHINSOKU.xlsm")
        ws = wb.Worksheets("NewDashboard")
        was_protected = ws.ProtectContents
        if was_protected:
            ws.Unprotect(Password="")
        for col, formula in COLUMN_FORMULAS.items():
            set_column(ws, col, formula)
        if was_protected:
            ws.Protect(
                Password="",
                UserInterfaceOnly=True,
                AllowFormattingCells=True,
                AllowSorting=True,
                AllowFiltering=True,
            )
        wb.Save()
        wb.Close()
    finally:
        xl.Quit()


if __name__ == "__main__":
    main()

