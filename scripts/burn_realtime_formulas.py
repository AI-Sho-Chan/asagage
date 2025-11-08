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
ROWS = 400

HN_TICKER = 'Ticker'
FIELD_NAME = '銘柄名称'
FIELD_LAST = '現在値'
FIELD_VWAP = '出来高加重平均'
FIELD_PREV = '前日終値'
FIELD_BID = '気配値(買)'
FIELD_ASK = '気配値(売)'
HN_MID = '気配値(中央)'
HN_ATR_N = 'ATR_n'
HN_JTH = 'J_th'
HN_TICKER_SRC = 'TickerSrc'
HN_SIG_STATUS = 'シグナル点灯'
HN_SIG_KIND = 'シグナル種別'
HN_GAP_PCT = '離脱距離率(%)'
HN_BID_LABEL = '気配値（買）'
HN_ASK_LABEL = '気配値（売）'
HN_MID_LABEL = '気配値（中央）'


def r1c(from_col: int, to_col: int) -> str:
    offset = from_col - to_col
    return 'RC' if offset == 0 else f'RC[{offset}]'


def main() -> None:
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

        # Unprotect if protected so COM can write formulas
        try:
            ws.Unprotect(Password="")
        except Exception:
            pass

        def set_col_formula(col: int, formula_r1c1: str, label: str) -> None:
            rng = ws.Range(ws.Cells(DATA_START, col), ws.Cells(DATA_START + ROWS, col))
            first = rng.Cells(1, 1)
            try:
                first.FormulaR1C1 = formula_r1c1
            except Exception as exc:
                print('ERR set', label, exc)
                return
            try:
                first.AutoFill(Destination=rng)
            except Exception as exc:
                print('WARN autofill', label, exc)

        last_col = ws.Cells(HEADER_ROW, ws.Columns.Count).End(1).Column
        t_col = 8
        if str(ws.Cells(HEADER_ROW, t_col).Value) != HN_TICKER:
            for c in range(1, last_col + 1):
                if str(ws.Cells(HEADER_ROW, c).Value) == HN_TICKER:
                    t_col = c
                    break

        ts_col = None
        for c in range(1, last_col + 1):
            if str(ws.Cells(HEADER_ROW, c).Value) == HN_TICKER_SRC:
                ts_col = c
                break
        if ts_col:
            set_col_formula(t_col, f'=IF({r1c(ts_col, t_col)}="","",{r1c(ts_col, t_col)})', HN_TICKER)

        name_c = t_col + 1
        jval_c = t_col + 2
        sig_st_c = t_col + 4
        sig_k_c = t_col + 5
        last_c = t_col + 6
        vwap_c = t_col + 7
        prev_c = t_col + 8
        bid_c = t_col + 9
        ask_c = t_col + 10
        mid_c = t_col + 11
        gap_bp_c = t_col + 12
        buck_c = t_col + 13
        act_c = t_col + 14

        set_col_formula(sig_st_c, '=""', HN_SIG_STATUS)
        set_col_formula(sig_k_c, '=""', HN_SIG_KIND)

        set_col_formula(name_c, f'=IF({r1c(t_col, name_c)}="","",IFERROR(RssMarket({r1c(t_col, name_c)},"{FIELD_NAME}"),""))', FIELD_NAME)
        set_col_formula(last_c, f'=IF({r1c(t_col, last_c)}="","",IFERROR(RssMarket({r1c(t_col, last_c)},"{FIELD_LAST}"),""))', FIELD_LAST)
        set_col_formula(vwap_c, f'=IF({r1c(t_col, vwap_c)}="","",IFERROR(RssMarket({r1c(t_col, vwap_c)},"{FIELD_VWAP}"),""))', FIELD_VWAP)
        set_col_formula(prev_c, f'=IF({r1c(t_col, prev_c)}="","",IFERROR(RssMarket({r1c(t_col, prev_c)},"{FIELD_PREV}"),""))', FIELD_PREV)
        set_col_formula(bid_c, f'=IF({r1c(t_col, bid_c)}="","",IFERROR(RssMarket({r1c(t_col, bid_c)},"{FIELD_BID}"),""))', FIELD_BID)
        set_col_formula(ask_c, f'=IF({r1c(t_col, ask_c)}="","",IFERROR(RssMarket({r1c(t_col, ask_c)},"{FIELD_ASK}"),""))', FIELD_ASK)
        set_col_formula(mid_c, f'=IF(OR({r1c(bid_c, mid_c)}="",{r1c(ask_c, mid_c)}=""),"",({r1c(bid_c, mid_c)}+{r1c(ask_c, mid_c)})/2)', HN_MID)

        atr_c = None
        for c in range(1, last_col + 1):
            if str(ws.Cells(HEADER_ROW, c).Value) == HN_ATR_N:
                atr_c = c
                break
        if atr_c:
            set_col_formula(jval_c, f'=IF(OR({r1c(last_c, jval_c)}="",{r1c(vwap_c, jval_c)}="",{r1c(atr_c, jval_c)}=0),"",({r1c(last_c, jval_c)}-{r1c(vwap_c, jval_c)})/{r1c(atr_c, jval_c)})', 'JValue')

        set_col_formula(gap_bp_c, f'=IF(OR({r1c(mid_c, gap_bp_c)}="",{r1c(prev_c, gap_bp_c)}=""),"",({r1c(mid_c, gap_bp_c)}-{r1c(prev_c, gap_bp_c)})/{r1c(prev_c, gap_bp_c)}*10000)', 'GapBp')

        # K 列（閾値乖離率%）= |J - J_th| / |J_th| * 100
        # jval は jval_c、K は jval_c + 1 とみなす
        k_c = jval_c + 1
        jth_c = None
        for c in range(1, last_col + 1):
            if str(ws.Cells(HEADER_ROW, c).Value) == HN_JTH:
                jth_c = c
                break
        if jth_c:
            set_col_formula(k_c, f'=IF(OR({r1c(jth_c, k_c)}="",{r1c(jval_c, k_c)}="",{r1c(jth_c, k_c)}=0),"",ABS({r1c(jval_c, k_c)}-{r1c(jth_c, k_c)})/ABS({r1c(jth_c, k_c)})*100)', 'JGapPct')

        set_col_formula(buck_c, f'=IF({r1c(gap_bp_c, buck_c)}="","",IF(ABS({r1c(gap_bp_c, buck_c)})>=120,">=120bp",IF(ABS({r1c(gap_bp_c, buck_c)})>=80,"80-120bp",IF(ABS({r1c(gap_bp_c, buck_c)})>=50,"50-80bp","<50bp"))))', 'GapBucket')
        set_col_formula(act_c, f'=IF({r1c(buck_c, act_c)}="","",IF({r1c(buck_c, act_c)}=">=120bp","j-cross only; TP-0.2; SL+0.2",IF({r1c(buck_c, act_c)}="80-120bp","Skip opposite; J_th+0.2",IF({r1c(buck_c, act_c)}="50-80bp","J_th+0.1","Baseline"))))', 'GapAction')

        # Re-apply in-VBA installer to add conditional formats and protect again
        try:
            excel.Run('AutoTrader.InstallRealtimeFormulas')
        except Exception:
            pass
        print('Burn completed.')
    finally:
        wb.Close(SaveChanges=True)
        excel.Quit()


if __name__ == '__main__':
    main()
