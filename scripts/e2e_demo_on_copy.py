import argparse
from pathlib import Path


def ensure_headers_and_dummy(wb, ticker: str, qty: int) -> None:
    ws = None
    try:
        ws = wb.Worksheets("NewDashboardV2")
    except Exception:
        ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "NewDashboardV2"
    headers = [
        "Ticker","Selected","J","J_th","EntryBuyPx","EntrySellPx","EntrySide","EntryStatus",
        "TP_price","SL_price","StopTrail","BestBid","BestAsk","PrevClose","VWAP","Gap_bp","CorrNKY","OrderQtyPlan"
    ]
    row = 5
    for i, h in enumerate(headers, start=1):
        ws.Cells(row, i).Value = h
    # dummy row 6
    ws.Cells(6, 1).Value = ticker
    ws.Cells(6, 2).Value = 1
    ws.Cells(6, 3).Value = 1.2
    ws.Cells(6, 4).Value = 1.0
    ws.Cells(6, 14).Value = 3000
    ws.Cells(6, 15).Value = 3010
    ws.Cells(6, 17).Value = 0.8
    ws.Cells(6, 18).Value = int(qty)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", required=True)
    ap.add_argument("--ticker", default="7203.T")
    ap.add_argument("--qty", type=int, default=100)
    args = ap.parse_args()

    import win32com.client  # type: ignore
    xl = win32com.client.DispatchEx("Excel.Application")
    try:
        xl.AutomationSecurity = 1  # msoAutomationSecurityLow
    except Exception:
        pass
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(Path(args.excel)))
        try:
            xl.Run(f"{wb.Name}!AutoTraderAdvanced.SetupNewDashboardV2")
        except Exception:
            pass
        ensure_headers_and_dummy(wb, args.ticker, args.qty)
        # Try macros first
        try:
            xl.Run(f"{wb.Name}!AutoTraderAdvanced.ApplyDynamicSignalsV2")
            xl.Run(f"{wb.Name}!AutoTraderAdvanced.PreplaceOrdersV2")
        except Exception:
            # Fallback: log pre-orders directly
            ws = wb.Worksheets("NewDashboardV2")
            # find columns by header row 5
            def col_of(name: str) -> int:
                for c in range(1,40):
                    if ws.Cells(5,c).Value == name:
                        return c
                return 0
            r = 6
            c_t = col_of("Ticker"); c_sel=col_of("Selected"); c_j=col_of("J"); c_jt=col_of("J_th"); c_v=col_of("VWAP"); c_p=col_of("PrevClose")
            t = str(ws.Cells(r,c_t).Value)
            sel = int(ws.Cells(r,c_sel).Value or 0)
            j = float(ws.Cells(r,c_j).Value or 0)
            jt = float(ws.Cells(r,c_jt).Value or 0)
            v = ws.Cells(r,c_v).Value or ws.Cells(r,c_p).Value
            try:
                v = float(v)
            except Exception:
                v = 0.0
            pre_frac = 0.5
            if sel == 1 and jt and abs(j) >= pre_frac*abs(jt) and v > 0:
                k = 0.001
                e_buy = v - k*abs(jt)*v
                e_sell = v + k*abs(jt)*v
                # Orders sheet
                try:
                    sh = wb.Worksheets("Orders")
                except Exception:
                    sh = wb.Worksheets.Add(After=ws); sh.Name = "Orders"
                    sh.Range("A1:K1").Value = ["ts","ticker","side","price","qty","mode","status","note","tp","sl","trail"]
                import datetime as dt
                ts = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                nr = sh.Cells(sh.Rows.Count,1).End(-4162).Row + 1  # xlUp = -4162
                sh.Cells(nr,1).Resize(1,11).Value = [ts,t,"BUY",e_buy,int(args.qty),"preplace","PENDING","V2","","",""]
                sh.Cells(nr+1,1).Resize(1,11).Value = [ts,t,"SELL",e_sell,int(args.qty),"preplace","PENDING","V2","","",""]
        wb.Save()
    finally:
        wb.Close(SaveChanges=True)
        xl.Quit()


if __name__ == "__main__":
    main()
