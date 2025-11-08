import argparse
import datetime as dt
import shutil
from pathlib import Path


def backup(path: Path) -> Path:
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = path.with_name(f"{path.stem}_backup_{ts}{path.suffix}")
    shutil.copy2(path, bak)
    return bak


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", default=r"C:/AI/asagake/ASAGAKE.xlsm")
    ap.add_argument("--ticker", default="7203.T")
    ap.add_argument("--qty", type=int, default=100)
    args = ap.parse_args()

    xlsm = Path(args.excel).resolve()
    backup(xlsm)

    import win32com.client  # type: ignore
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(xlsm))
        try:
            vbproj = wb.VBProject
        except Exception:
            raise SystemExit("This Excel is not trusted for VBA project access. Enable 'Trust access to the VBA project object model'.")

        # Install Advanced module via Import (safer)
        try:
            vbcomp = vbproj.VBComponents("AutoTraderAdvanced")
            vbproj.VBComponents.Remove(vbcomp)
        except Exception:
            pass
        vbproj.VBComponents.Import(str(Path("excel/AutoTraderAdvanced.bas").resolve()))
        wb.Save()

        # Prepare NewDashboardV2 with one dummy row
        try:
            ws = wb.Worksheets("NewDashboardV2")
        except Exception:
            wb.Application.Run(f"{wb.Name}!AutoTraderAdvanced.SetupNewDashboardV2")
            ws = wb.Worksheets("NewDashboardV2")

        # Headers already set by Setup; fill row 6
        ws.Cells(6, 1).Value = args.ticker  # Ticker
        ws.Cells(6, 2).Value = 1           # Selected
        ws.Cells(6, 3).Value = 1.20        # J
        ws.Cells(6, 4).Value = 1.00        # J_th (base)
        ws.Cells(6, 17).Value = 0.80       # CorrNKY
        ws.Cells(6, 18).Value = int(args.qty)

        # Set parameters on sheet (bias/gap/coefficients)
        ws.Range("B41").Value = 0         # Market bias bp
        ws.Range("B42").Value = 0.10      # Bias slope per 100bp
        ws.Range("B43").Value = 0.20      # Gap slope per 1%
        ws.Range("B44").Value = 3.0       # Hard ban gap %
        ws.Range("B45").Value = 5         # No-trade minutes from open
        ws.Range("B46").Value = 0.15      # TP per J_excess
        ws.Range("B47").Value = 0.10      # SL per J_excess
        ws.Range("B48").Value = 0.10      # Trail per J_excess
        ws.Range("B49").Value = 1000000   # Budget per ticker
        ws.Range("B50").Value = 0.5       # Preplace fraction

        # Install formulas and compute dynamics（RSSが利用可能な環境で有効）
        OnError = False
        try:
            wb.Application.Run(f"{wb.Name}!AutoTraderAdvanced.InstallRealtimeFormulasV2")
        except Exception:
            # RSSが未接続でも後段の式/ログは動作するため継続
            pass
        wb.Application.Run(f"{wb.Name}!AutoTraderAdvanced.ApplyDynamicSignalsV2")

        # Attempt pre-place (will try LIVE order via RSS, fallback logs to Orders)
        wb.Application.Run(f"{wb.Name}!AutoTraderAdvanced.PreplaceOrdersV2")

        wb.Save()
        wb.Close(SaveChanges=True)
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
