from __future__ import annotations

from pathlib import Path
import sys

try:
    import win32com.client  # type: ignore
except Exception as e:
    print("PYWIN32_IMPORT_ERROR", e)
    sys.exit(1)

WB_PATH = Path("C:/AI/asagake/SHINSOKU.xlsm")
SHEET = "NewDashboard"
ROW = 6
CHECK_COLS = {
    "I": True,   # should contain RssMarket
    "J": False,  # calc column
    "K": False,  # calc column
    "N": True,   # should contain RssMarket
    "O": True,   # should contain RssMarket
    "P": True,   # should contain RssMarket
    "Q": True,   # numeric code ok
    "R": True,   # numeric code ok
}


def main() -> int:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        excel.AutomationSecurity = 1
    except Exception:
        pass

    wb = excel.Workbooks.Open(str(WB_PATH))
    try:
        ws = wb.Worksheets(SHEET)
        failures = []
        for col, expect_rss in CHECK_COLS.items():
            c = ws.Range(f"{col}{ROW}")
            has_formula = bool(c.HasFormula)
            formula = str(c.Formula)
            if not has_formula:
                failures.append(f"{col}{ROW}: missing formula")
                continue
            if expect_rss and ("RssMarket" not in formula and "RssMarketHeader" not in formula):
                failures.append(f"{col}{ROW}: formula not RSS: {formula}")

        # Try a light refresh via Calculate; ignore errors
        try:
            excel.CalculateFull()
        except Exception:
            pass

        if failures:
            print("SMOKE_CHECK_FAIL")
            for f in failures:
                print(" -", f)
            return 1
        print("SMOKE_CHECK_OK: formulas present at I6 et al.")
        return 0
    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()


if __name__ == "__main__":
    raise SystemExit(main())

