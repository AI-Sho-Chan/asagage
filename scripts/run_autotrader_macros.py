from pathlib import Path
import time
import sys

try:
    import win32com.client  # type: ignore
except Exception as e:
    print("PYWIN32_IMPORT_ERROR", e)
    sys.exit(1)

WB_PATH = Path("C:/AI/asagake/SHINSOKU.xlsm")


def main():
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    try:
        # 1 = msoAutomationSecurityLow (allow macros)
        excel.AutomationSecurity = 1
    except Exception:
        pass
    wb = excel.Workbooks.Open(str(WB_PATH))
    try:
        print("Running AutoTrader.ResetDashboardHeaders ...")
        excel.Run("AutoTrader.ResetDashboardHeaders")
        time.sleep(0.5)
        print("Running AutoTrader.ButtonLoadCandidates ...")
        excel.Run("AutoTrader.ButtonLoadCandidates")
        time.sleep(0.2)
        print("Running AutoTrader.ButtonPushCandidates ...")
        excel.Run("AutoTrader.ButtonPushCandidates")
        time.sleep(0.5)
        ws = wb.Worksheets("NewDashboard")
        # Dump a few key cells: H6..N6 and some names/last/vwap
        def f(cell):
            try:
                return ws.Range(cell).Formula
            except Exception:
                return None
        print("H6:", f("H6"))
        print("I6:", f("I6"))
        print("J6:", f("J6"))
        print("K6:", f("K6"))
        print("L6:", f("L6"))
        print("M6:", f("M6"))
        print("N6:", f("N6"))
    finally:
        wb.Close(SaveChanges=True)
        excel.Quit()


if __name__ == "__main__":
    main()
