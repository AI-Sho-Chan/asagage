from pathlib import Path
import sys

try:
    import win32com.client  # type: ignore
except Exception as e:
    print("PYWIN32_IMPORT_ERROR", e)
    sys.exit(1)

WB_PATH = Path("C:/AI/asagake/SHINSOKU.xlsm")
SHEET_MAIN = "NewDashboard"
HEADER_ROW = 5
DATA_START = 6
ROWS = 400


def delete_sheet_if_exists(app, wb, name: str) -> bool:
    try:
        sh = wb.Worksheets(name)
    except Exception:
        return False
    app.DisplayAlerts = False
    try:
        sh.Delete()
    finally:
        app.DisplayAlerts = True
    return True


def clear_columns(ws, cols):
    for c in cols:
        ws.Cells(HEADER_ROW, c).Value = ""
        ws.Range(ws.Cells(DATA_START, c), ws.Cells(DATA_START + ROWS, c)).ClearContents()


def main():
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(str(WB_PATH))
    try:
        removed = delete_sheet_if_exists(excel, wb, "Dashboard")
        if removed:
            print("Removed sheet: Dashboard")
        ws = wb.Worksheets(SHEET_MAIN)
        # Identify duplicate/unneeded columns beyond PlanTag (AR=44)
        to_clear = []
        for c in range(45, 61):  # AS(45) .. BI(60)
            v = ws.Cells(HEADER_ROW, c).Value
            if v in (None, ""):
                to_clear.append(c)
                continue
            if str(v) in ("PreOpenMid", "LiveGapBp", "Message", "DynamicQty"):
                to_clear.append(c)
        clear_columns(ws, to_clear)
        print("Cleared columns:", to_clear)
        wb.Save()
    finally:
        wb.Close(SaveChanges=True)
        excel.Quit()


if __name__ == "__main__":
    main()

