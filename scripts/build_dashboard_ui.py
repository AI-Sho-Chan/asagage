import argparse
import csv
from pathlib import Path


def build_ui(copy_path: Path) -> None:
    import win32com.client  # type: ignore
    xl = win32com.client.DispatchEx("Excel.Application")
    try:
        # 1 = msoAutomationSecurityLow
        xl.AutomationSecurity = 1  # type: ignore[attr-defined]
    except Exception:
        pass
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(copy_path))
        # Import latest VBA module (ASCII-only)
        v = wb.VBProject
        modules = [
            ("AutoTraderAdvanced", "excel/AutoTraderAdvanced.bas"),
            ("cDashboardWatcher", "excel/cDashboardWatcher.cls"),
        ]
        for mod_name, rel_path in modules:
            try:
                vb = v.VBComponents(mod_name)
                v.VBComponents.Remove(vb)
            except Exception:
                pass
            v.VBComponents.Import(str(Path(rel_path).resolve()))

        try:
            ws = wb.Worksheets("NewDashboardV2")
        except Exception:
            ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            ws.Name = "NewDashboardV2"

        # Parameters row 1 (labels) / row 2 (values)
        labels = [
            "NKY_Code","NKY_Last","NKY_ChgPct","Bias_bp","BiasSlope","GapSlope",
            "GapBanPct","NoTradeMin","TP_per_J","SL_per_J","Trail_per_J","CorrSlope",
        ]
        defaults = ["N225","",0.0,0.0,0.10,0.20,3.0,5,0.15,0.10,0.10,0.05]
        for i,(h, val) in enumerate(zip(labels, defaults), start=1):
            ws.Cells(1,i).Value = h
            ws.Cells(2,i).Value = val
        # NKY auto (requires RSS add-in). 現在値/騰落率→Bias_bp(bp)
        try:
            ws.Cells(2,2).FormulaR1C1Local = '=IF(RC[-1]="", "", IFERROR(RssIndexMarket(RC[-1],"現在値"),""))'
            ws.Cells(2,3).FormulaR1C1Local = '=IF(RC[-2]="", "", IFERROR(RssIndexMarket(RC[-2],"騰落率"),""))'
            ws.Cells(2,4).FormulaR1C1 = "=(RC[-1])*100"
        except Exception:
            pass

        # Clear old banner (if any) and place new one at N1:U2
        try:
            old = ws.Range("B2:J3")
            if old.MergeCells:
                old.UnMerge()
            old.Clear()
        except Exception:
            pass
        banner = ws.Range("N1:U2"); banner.Merge()
        banner.HorizontalAlignment = -4108; banner.VerticalAlignment = -4108
        banner.Font.Bold = True; banner.Font.Size = 18
        banner.Interior.Color = 15132390  # RGB(230,230,230)
        banner.Value = "IDLE"; banner.Name = "RunStatusV2"

        # Remove old buttons
        try:
            for shp in list(ws.Shapes):
                if shp.Name.startswith("btn_"): shp.Delete()
        except Exception:
            pass

        def add_btn(name: str, caption: str, r: int, c: int, macro: str) -> None:
            cell = ws.Cells(r,c)
            left, top = cell.Left, cell.Top
            width = ws.Range(ws.Cells(r,c), ws.Cells(r,c+3)).Width - 6
            height = ws.Rows(r).Height - 2
            shp = ws.Shapes.AddShape(1, left, top, width, height)  # 1 = msoShapeRectangle
            try:
                shp.Name = name
            except Exception:
                pass
            shp.TextFrame.Characters().Text = caption
            shp.TextFrame.HorizontalAlignment = -4108
            shp.TextFrame.VerticalAlignment = -4108
            shp.OnAction = macro

        # Buttons: swap groups (Live left, Demo right) and move to row3 right area
        add_btn("btn_live_start","\u672c\u756a\u53d6\u5f15\u958b\u59cb",3,14,"AutoTraderAdvanced.StartLiveV2")
        add_btn("btn_live_stop","\u672c\u756a\u53d6\u5f15\u505c\u6b62",3,18,"AutoTraderAdvanced.StopLiveV2")
        add_btn("btn_demo_start","\u30c7\u30e2\u53d6\u5f15\u958b\u59cb",3,22,"AutoTraderAdvanced.StartDemoV2")
        add_btn("btn_demo_stop","\u30c7\u30e2\u53d6\u5f15\u505c\u6b62",3,26,"AutoTraderAdvanced.StopDemoV2")
        add_btn("btn_import","\u5019\u88dc\u9298\u67c4\u53d6\u8fbc",3,30,"AutoTraderAdvanced.ImportCandidatesV2")

        # Row4 JP labels + notes
        jp = [
            "\u8a3c\u5238\u30b3\u30fc\u30c9","\u9298\u67c4\u540d","J_th\u30d9\u30fc\u30b9","J_th(\u88dc\u6b63\u5f8c)","J\u5024","\u73fe\u5728\u5024","\u524d\u65e5\u7d42\u5024","VWAP","\u767a\u6ce8\u4e88\u5b9a\u6570\u91cf","\u76e3\u8996ON/OFF",
            "\u8cb7\u6307\u5024","\u58f2\u6307\u5024","\u30b5\u30a4\u30c9","\u767a\u6ce8\u72b6\u6cc1","\u5229\u78ba\u6307\u5024","\u640d\u5207\u6307\u5024",
            "\u30c8\u30ec\u30fc\u30ea\u30f3\u30b0","\u6c7a\u6e08\u72b6\u6cc1","\u6700\u826f\u8cb7\u6c17\u914d\u5024","\u6700\u826f\u58f2\u6c17\u914d\u5024","\u30ae\u30e3\u30c3\u30d7(bp)","\u76f8\u95a2(NKY)",
            "Bias\u4fc2\u6570","\u30ae\u30e3\u30c3\u30d7\u4fc2\u6570","Corr\u4fc2\u6570","TP/J(\u9ad8\u6821)","SL/J(\u9ad8\u6821)","Trail/J(\u9ad8\u6821)","TP/J(\u5b9f\u52d5)","SL/J(\u5b9f\u52d5)","Trail/J(\u5b9f\u52d5)","Vol\u30bf\u30b0"
        ]
        for i,text in enumerate(jp, start=1):
            ws.Cells(4,i).Value = text
            try:
                ws.Cells(4,i).AddComment("\u3053\u306e\u5217\u306e\u610f\u5473\u3068\u6d3b\u7528\u65b9\u6cd5\u3092\u8aac\u660e\u3057\u307e\u3059\u3002")
            except Exception:
                pass

        # Row5 headers (order)
        hdr = ["Ticker","Name","J_th_base","J_th","J","Last","PrevClose","VWAP","OrderQtyPlan","Selected",
                "EntryBuyPx","EntrySellPx","EntrySide","EntryStatus","TP_price","SL_price",
                "StopTrail","SettleStatus","BestBid","BestAsk","Gap_bp","CorrNKY",
                "BiasSlope_row","GapSlope_row","CorrSlope_row","TP_per_J_row","SL_per_J_row","Trail_per_J_row",
                "TP_per_J_eff","SL_per_J_eff","Trail_per_J_eff","VolatilityTag"]
        for i,text in enumerate(hdr, start=1):
            ws.Cells(5,i).Value = text
        # Append forward metrics columns on the right
        extras_hdr = ["ForwardPfEff","WinCiLow","ForwardTrades","ExpBp"]
        extras_jp  = ["PF(フォワード)","勝率CI下限","トレード回数","期待bp"]
        # find last header col
        last_c = 1
        for c in range(1,60):
            if ws.Cells(5,c).Value:
                last_c = c
        for idx, name in enumerate(extras_hdr, start=1):
            ws.Cells(5, last_c+idx).Value = name
            ws.Cells(4, last_c+idx).Value = extras_jp[idx-1]
            try:
                ws.Cells(4, last_c+idx).AddComment("ナイト/週次で推定された指標。監視と配分に利用")
            except Exception:
                pass

        # Import candidates (weekly or nightly)
        root = Path("C:/AI/asagake/output/excel")
        cand = root / "candidates_nextday.csv"
        if not cand.exists():
            ws_f = sorted(root.glob("weekly_candidates_*.csv"))
            if ws_f:
                cand = ws_f[-1]
        tickers = []
        if cand.exists():
            with cand.open("r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    t = row.get("Ticker") or row.get("code")
                    if not t: continue
                    tickers.append(t)
        r0 = 6
        for i,t in enumerate(tickers, start=0):
            ws.Cells(r0+i,1).Value = t
            ws.Cells(r0+i,8).Value = 1

        # RSS formulas for first row (Name/PrevClose/VWAP/BestBid/BestAsk/Gap)
        def col_of(name: str) -> int:
            for c in range(1,60):
                if ws.Cells(5,c).Value == name:
                    return c
            return 0
        c_t = col_of("Ticker"); c_last=col_of("Last"); c_prev=col_of("PrevClose"); c_vwap=col_of("VWAP"); c_bid=col_of("BestBid"); c_ask=col_of("BestAsk"); c_name=col_of("Name"); c_gap=col_of("Gap_bp"); c_qty=col_of("OrderQtyPlan")
        c_jtb = col_of("J_th_base")
        if c_jtb:
            try:
                if not ws.Cells(6,c_jtb).Value:
                    ws.Cells(6,c_jtb).Value = 1.0
            except Exception:
                pass
        if c_last:
            ws.Cells(6,c_last).FormulaLocal = f'=IF(RC[{c_t-c_last}]="","",IFERROR(RssMarket(SUBSTITUTE(RC[{c_t-c_last}],".T",""),"\u73fe\u5728\u5024"),""))'
        if c_prev:
            ws.Cells(6,c_prev).FormulaLocal = f'=IF(RC[{c_t-c_prev}]="","",IFERROR(RssMarket(SUBSTITUTE(RC[{c_t-c_prev}],".T",""),"\u524d\u65e5\u7d42\u5024"),""))'
        if c_vwap:
            ws.Cells(6,c_vwap).FormulaLocal = f'=IF(RC[{c_t-c_vwap}]="","",IFERROR(RssMarket(SUBSTITUTE(RC[{c_t-c_vwap}],".T",""),"VWAP"),""))'
        if c_bid:
            ws.Cells(6,c_bid).FormulaLocal = f'=IF(RC[{c_t-c_bid}]="","",IFERROR(RssMarket(SUBSTITUTE(RC[{c_t-c_bid}],".T",""),"\u6700\u826f\u8cb7\u6c17\u914d\u5024"),""))'
        if c_ask:
            ws.Cells(6,c_ask).FormulaLocal = f'=IF(RC[{c_t-c_ask}]="","",IFERROR(RssMarket(SUBSTITUTE(RC[{c_t-c_ask}],".T",""),"\u6700\u826f\u58f2\u6c17\u914d\u5024"),""))'
        if c_name:
            ws.Cells(6,c_name).FormulaLocal = f'=IF(RC[{c_t-c_name}]="","",IFERROR(RssMarket(SUBSTITUTE(RC[{c_t-c_name}],".T",""),"\u9298\u67c4\u540d"),""))'
        if c_gap:
            ws.Cells(6,c_gap).FormulaLocal = f'=IF(OR(RC[{c_prev-c_gap}]="",RC[{c_vwap-c_gap}]=""),"",(RC[{c_vwap-c_gap}]-RC[{c_prev-c_gap}])/RC[{c_prev-c_gap}]*10000)'
        if c_qty and (c_last or c_prev):
            try:
                ref = c_last if c_last else c_prev
                ws.Range(ws.Cells(6,c_qty), ws.Cells(605,c_qty)).FormulaR1C1 = (
                    f'=IF(OR(RC[{ref-c_qty}]="",RC[{ref-c_qty}]=0),"",MAX(0,INT(R2C13/RC[{ref-c_qty}]/R2C14)*R2C14))'
                )
            except Exception:
                ref = c_last if c_last else c_prev
                ws.Range(ws.Cells(6,c_qty), ws.Cells(605,c_qty)).FormulaR1C1Local = (
                    f'=IF(OR(RC[{ref-c_qty}]="",RC[{ref-c_qty}]=0),"",MAX(0,INT(R2C13/RC[{ref-c_qty}]/R2C14)*R2C14))'
                )

        # Try to run macros if available; continue even if disabled
        try:
            wb.Application.Run(f"{wb.Name}!AutoTraderAdvanced.ApplyDynamicSignalsV2")
            wb.Application.Run(f"{wb.Name}!AutoTraderAdvanced.PreplaceOrdersV2")
        except Exception:
            pass
        wb.Save()
    finally:
        wb.Close(SaveChanges=True)
        xl.Quit()


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", required=True)
    args = ap.parse_args()
    build_ui(Path(args.excel))


if __name__ == "__main__":
    main()


