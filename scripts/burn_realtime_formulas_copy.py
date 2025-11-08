from __future__ import annotations

from pathlib import Path
import sys

try:
    import win32com.client  # type: ignore
except Exception as e:
    print("PYWIN32_IMPORT_ERROR", e)
    sys.exit(1)

WB_PATH = Path("C:/AI/asagake/SHINSOKU.xlsm")
SHEET_NAME = "NewDashboard"
HEADER_ROW = 5
DATA_START = 6
ROWS = 400

# Header labels must match the sheet exactly
HN_TICKER = "Ticker"
HN_NAME = "\u9298\u67c4\u540d"  # 銘柄名
HN_JVAL = "\u73fe\u5728\u306eJ\u5024"  # 現在のJ値
HN_GAP_PCT = "\u95be\u5024\u4e56\u96e2\u7387(%)"  # 乖離率(%)
HN_SIG_STATUS = "\u30b7\u30b0\u30ca\u30eb\u70b9\u706f"  # シグナル点灯
HN_SIG_KIND = "\u30b7\u30b0\u30ca\u30eb\u7a2e\u5225"   # シグナル種別
HN_LAST = "\u73fe\u5728\u5024"  # 現在値
HN_VWAP = "\u51fa\u6765\u9ad8\u52a0\u91cd\u5e73\u5747"  # 出来高加重平均
HN_PREV = "\u524d\u65e5\u7d42\u5024"  # 前日終値
HN_BID = "\u6c17\u914d\u5024(\u8cb7)"  # 気配値(買)
HN_ASK = "\u6c17\u914d\u5024(\u58f2)"  # 気配値(売)
HN_MID = "\u6c17\u914d\u5024(\u4e2d\u592e)"  # 気配値(中央)
HN_GAP_BP = "\u30e9\u30a4\u30d6\u30ae\u30e3\u30c3\u30d7(bp)"  # ライブギャップ(bp)
HN_GAP_BUCKET = "\u30e9\u30a4\u30d6\u30ae\u30e3\u30c3\u30d7\u5e2f"  # ライブギャップ帯
HN_ACTION = "\u30e9\u30a4\u30d6\u30a2\u30af\u30b7\u30e7\u30f3"  # ライブアクション
HN_ATR_N = "ATR_n"
HN_JTH = "J_th"
HN_TICKER_SRC = "TickerSrc"


def r1c(from_col: int, to_col: int) -> str:
    off = from_col - to_col
    return "RC" if off == 0 else f"RC[{off}]"


def main():
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        excel.AutomationSecurity = 1
    except Exception:
        pass
    wb = excel.Workbooks.Open(str(WB_PATH))
    try:
        ws = wb.Worksheets(SHEET_NAME)

        def set_col_formula_by_index(c: int, formula_r1c1: str, label: str):
            rng = ws.Range(ws.Cells(DATA_START, c), ws.Cells(DATA_START + ROWS, c))
            first = rng.Cells(1, 1)
            try:
                first.FormulaR1C1 = formula_r1c1
            except Exception as e:
                print("ERR set", label, e)
                return
            try:
                first.AutoFill(Destination=rng)
            except Exception as e:
        # locate anchor columns by position
        t_col = ws.Cells(HEADER_ROW, ws.Columns.Count).End(1).Column  # temp reuse
        # Correct anchor: find Ticker at col >=8
        t_col = 8  # default expected
        if str(ws.Cells(HEADER_ROW, t_col).Value) != HN_TICKER:
            # Fallback: search row for Ticker
            last_col = ws.Cells(HEADER_ROW, ws.Columns.Count).End(1).Column
            for c in range(1, last_col + 1):
                if str(ws.Cells(HEADER_ROW, c).Value) == HN_TICKER:
                    t_col = c
                    break

        # Candidate source (if present)
        last_col = ws.Cells(HEADER_ROW, ws.Columns.Count).End(1).Column
        ts_col = None
        for c in range(1, last_col + 1):
            if str(ws.Cells(HEADER_ROW, c).Value) == HN_TICKER_SRC:
                ts_col = c
                break

        if ts_col:
            set_col_formula_by_index(t_col, f'=IF({r1c(ts_col, t_col)}="","",{r1c(ts_col, t_col)})', HN_TICKER)

        # realtime offsets relative to ticker
        name_c = t_col + 1
        jval_c = t_col + 2
        gap_pct_c = t_col + 3
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

        # Static blanks for signal columns
        set_col_formula_by_index(sig_st_c, '=""', HN_SIG_STATUS)
        set_col_formula_by_index(sig_k_c, '=""', HN_SIG_KIND)

        # Rss formulas using numeric codes
        set_col_formula_by_index(name_c, f'=IF({r1c(t_col, name_c)}="","",IFERROR(RssMarket({r1c(t_col, name_c)},3),""))', HN_NAME)
        set_col_formula_by_index(last_c, f'=IF({r1c(t_col, last_c)}="","",IFERROR(RssMarket({r1c(t_col, last_c)},8),""))', HN_LAST)
        set_col_formula_by_index(vwap_c, f'=IF({r1c(t_col, vwap_c)}="","",IFERROR(RssMarket({r1c(t_col, vwap_c)},28),""))', HN_VWAP)
        set_col_formula_by_index(prev_c, f'=IF({r1c(t_col, prev_c)}="","",IFERROR(RssMarket({r1c(t_col, prev_c)},15),""))', HN_PREV)
        set_col_formula_by_index(bid_c, f'=IF({r1c(t_col, bid_c)}="","",IFERROR(RssMarket({r1c(t_col, bid_c)},56),""))', HN_BID)
        set_col_formula_by_index(ask_c, f'=IF({r1c(t_col, ask_c)}="","",IFERROR(RssMarket({r1c(t_col, ask_c)},55),""))', HN_ASK)
        set_col_formula_by_index(mid_c, f'=IF(OR({r1c(bid_c, mid_c)}="",{r1c(ask_c, mid_c)}=""),"",({r1c(bid_c, mid_c)}+{r1c(ask_c, mid_c)})/2)', HN_MID)
        set_col_formula_by_index(jval_c, f'=IF(OR({r1c(last_c, jval_c)}="",{r1c(vwap_c, jval_c)}="",{r1c(ws.Cells(HEADER_ROW, ws.Columns.Count).End(1).Column, jval_c)}=0),"",({r1c(last_c, jval_c)}-{r1c(vwap_c, jval_c)})/{r1c(t_col+ (HN_ATR_N and 0), jval_c)})', HN_JVAL)
        # The above ATR reference line is complex; replace with explicit by searching column label ATR_n
        # Recompute ATR reference properly
        atr_c = None
        for c in range(1, last_col + 1):
            if str(ws.Cells(HEADER_ROW, c).Value) == HN_ATR_N:
                atr_c = c
                break
        if atr_c:
            set_col_formula_by_index(jval_c, f'=IF(OR({r1c(last_c, jval_c)}="",{r1c(vwap_c, jval_c)}="",{r1c(atr_c, jval_c)}=0),"",({r1c(last_c, jval_c)}-{r1c(vwap_c, jval_c)})/{r1c(atr_c, jval_c)})', HN_JVAL)

        set_col_formula_by_index(gap_bp_c, f'=IF(OR({r1c(mid_c, gap_bp_c)}="",{r1c(prev_c, gap_bp_c)}=""),"",({r1c(mid_c, gap_bp_c)}-{r1c(prev_c, gap_bp_c)})/{r1c(prev_c, gap_bp_c)}*10000)', HN_GAP_BP)
        # Gap% bucket + action
        set_col_formula_by_index(gap_pct_c, f'=IF(OR({r1c(jval_c, gap_pct_c)}="",{r1c(t_col+ (0), gap_pct_c)}="",{r1c(t_col+(0), gap_pct_c)}=0),"",MAX(0,(ABS({r1c(t_col+(0), gap_pct_c)})-ABS({r1c(jval_c, gap_pct_c)}))/ABS({r1c(t_col+(0), gap_pct_c)})*100))', HN_GAP_PCT)
        # Replace the placeholder with actual J_th column index
        jth_c = None
        for c in range(1, last_col + 1):
            if str(ws.Cells(HEADER_ROW, c).Value) == HN_JTH:
                jth_c = c
                break
        if jth_c:
            set_col_formula_by_index(gap_pct_c, f'=IF(OR({r1c(jth_c, gap_pct_c)}="",{r1c(jval_c, gap_pct_c)}="",{r1c(jth_c, gap_pct_c)}=0),"",MAX(0,(ABS({r1c(jth_c, gap_pct_c)})-ABS({r1c(jval_c, gap_pct_c)}))/ABS({r1c(jth_c, gap_pct_c)})*100))', HN_GAP_PCT)
        set_col_formula_by_index(buck_c, f'=IF({r1c(gap_bp_c, buck_c)}="","",IF(ABS({r1c(gap_bp_c, buck_c)})>=120,">=120bp",IF(ABS({r1c(gap_bp_c, buck_c)})>=80,"80-120bp",IF(ABS({r1c(gap_bp_c, buck_c)})>=50,"50-80bp","<50bp"))))', HN_GAP_BUCKET)
        set_col_formula_by_index(act_c, f'=IF({r1c(buck_c, act_c)}="","",IF({r1c(buck_c, act_c)}=">=120bp","j-cross only; TP-0.2; SL+0.2",IF({r1c(buck_c, act_c)}="80-120bp","Skip opposite; J_th+0.2",IF({r1c(buck_c, act_c)}="50-80bp","J_th+0.1","Baseline"))))', HN_ACTION)

        print("Burn completed.")
    finally:
        wb.Close(SaveChanges=True)
        excel.Quit()


if __name__ == "__main__":
    main()
