from __future__ import annotations

from pathlib import Path
import sys

try:
    import win32com.client  # type: ignore
except Exception as e:
    print('PYWIN32_IMPORT_ERROR', e)
    sys.exit(1)

WB_PATH = Path('C:/AI/asagake/SHINSOKU.xlsm')
SHEET_NAME = 'NewDashboard'
HEADER_ROW = 5
DATA_START = 6
ROWS = 500


def jp(*codes: int) -> str:
    return ''.join(chr(c) for c in codes)


FIELD_NAME = jp(0x9298, 0x67C4, 0x540D, 0x79F0)  # 銘柄名称
FIELD_LAST = jp(0x73FE, 0x5728, 0x5024)          # 現在値
FIELD_VWAP = jp(0x51FA, 0x6765, 0x9AD8, 0x52A0, 0x91CD, 0x5E73, 0x5747)  # 出来高加重平均
FIELD_PREV = jp(0x524D, 0x65E5, 0x7D42, 0x5024)  # 前日終値
FIELD_BID = jp(0x6C17, 0x914D, 0x5024, 0xFF08, 0x8CB7, 0xFF09)  # 気配値(買)
FIELD_ASK = jp(0x6C17, 0x914D, 0x5024, 0xFF08, 0x58F2, 0xFF09)  # 気配値(売)


def r1c(from_col: int, to_col: int) -> str:
    off = from_col - to_col
    return 'RC' if off == 0 else f'RC[{off}]'


def set_col_formula(ws, start_row: int, rows: int, col: int, formula_r1c1: str) -> None:
    rng = ws.Range(ws.Cells(start_row, col), ws.Cells(start_row + rows, col))
    first = rng.Cells(1, 1)
    first.FormulaR1C1 = formula_r1c1
    try:
        first.AutoFill(Destination=rng)
    except Exception:
        pass


def main() -> int:
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        excel.AutomationSecurity = 1
    except Exception:
        pass

    wb = excel.Workbooks.Open(str(WB_PATH))
    try:
        ws = wb.Worksheets(SHEET_NAME)
        try:
            ws.Unprotect(Password='')
        except Exception:
            pass

        # Find Ticker header column (default H=8)
        t_col = 8
        last_col = ws.Cells(HEADER_ROW, ws.Columns.Count).End(1).Column
        if str(ws.Cells(HEADER_ROW, t_col).Value) != 'Ticker':
            for c in range(1, last_col + 1):
                if str(ws.Cells(HEADER_ROW, c).Value) == 'Ticker':
                    t_col = c
                    break

        # Compute column indices relative to Ticker
        name_c = t_col + 1  # I
        jval_c = t_col + 2  # J
        sig_st_c = t_col + 4  # L
        sig_k_c = t_col + 5   # M
        last_c = t_col + 6    # N
        vwap_c = t_col + 7    # O
        prev_c = t_col + 8    # P
        bid_c = t_col + 9     # Q
        ask_c = t_col + 10    # R
        mid_c = t_col + 11    # S
        gap_bp_c = t_col + 12 # T

        # Clear L/M via empty string and set downstream columns
        set_col_formula(ws, DATA_START, ROWS, sig_st_c, '=""')
        set_col_formula(ws, DATA_START, ROWS, sig_k_c, '=""')

        # RSS fields
        set_col_formula(ws, DATA_START, ROWS, name_c, f'=IF({r1c(t_col, name_c)}="","",IFERROR(RssMarket({r1c(t_col, name_c)},"{FIELD_NAME}"),""))')
        set_col_formula(ws, DATA_START, ROWS, last_c, f'=IF({r1c(t_col, last_c)}="","",IFERROR(RssMarket({r1c(t_col, last_c)},"{FIELD_LAST}"),""))')
        set_col_formula(ws, DATA_START, ROWS, vwap_c, f'=IF({r1c(t_col, vwap_c)}="","",IFERROR(RssMarket({r1c(t_col, vwap_c)},"{FIELD_VWAP}"),""))')
        set_col_formula(ws, DATA_START, ROWS, prev_c, f'=IF({r1c(t_col, prev_c)}="","",IFERROR(RssMarket({r1c(t_col, prev_c)},"{FIELD_PREV}"),""))')
        set_col_formula(ws, DATA_START, ROWS, bid_c, f'=IF({r1c(t_col, bid_c)}="","",IFERROR(RssMarket({r1c(t_col, bid_c)},"{FIELD_BID}"),""))')
        set_col_formula(ws, DATA_START, ROWS, ask_c, f'=IF({r1c(t_col, ask_c)}="","",IFERROR(RssMarket({r1c(t_col, ask_c)},"{FIELD_ASK}"),""))')

        # Derived columns
        set_col_formula(ws, DATA_START, ROWS, mid_c, f'=IF(OR({r1c(bid_c, mid_c)}="",{r1c(ask_c, mid_c)}=""),"",({r1c(bid_c, mid_c)}+{r1c(ask_c, mid_c)})/2)')
        set_col_formula(ws, DATA_START, ROWS, gap_bp_c, f'=IF(OR({r1c(mid_c, gap_bp_c)}="",{r1c(prev_c, gap_bp_c)}=""),"",({r1c(mid_c, gap_bp_c)}-{r1c(prev_c, gap_bp_c)})/{r1c(prev_c, gap_bp_c)}*10000)')

        try:
            excel.Run('AutoTrader.InstallRealtimeFormulas')
        except Exception:
            pass

        wb.Save()
        print('Install completed.')
        return 0
    finally:
        wb.Close(SaveChanges=True)
        excel.Quit()


if __name__ == '__main__':
    raise SystemExit(main())

