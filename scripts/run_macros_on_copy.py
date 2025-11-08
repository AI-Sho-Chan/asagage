import argparse
from pathlib import Path


def run_macros(copy_path: Path) -> None:
    import win32com.client  # type: ignore
    xl = win32com.client.DispatchEx("Excel.Application")
    # 1 = msoAutomationSecurityLow
    try:
        xl.AutomationSecurity = 1  # type: ignore[attr-defined]
    except Exception:
        pass
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(copy_path))
        try:
            xl.Run(f"{wb.Name}!AutoTraderAdvanced.SetupNewDashboardV2")
        except Exception:
            # sheet may already exist
            pass
        try:
            xl.Run(f"{wb.Name}!AutoTraderAdvanced.InstallRealtimeFormulasV2")
        except Exception:
            pass
        xl.Run(f"{wb.Name}!AutoTraderAdvanced.ApplyDynamicSignalsV2")
        xl.Run(f"{wb.Name}!AutoTraderAdvanced.PreplaceOrdersV2")
        wb.Save()
        wb.Close(SaveChanges=True)
    finally:
        xl.Quit()


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", required=True, help="Path to the copy workbook (.xlsm)")
    args = ap.parse_args()
    run_macros(Path(args.excel))


if __name__ == "__main__":
    main()

