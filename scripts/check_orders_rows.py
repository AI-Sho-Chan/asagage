import argparse
from pathlib import Path


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", required=True)
    args = ap.parse_args()

    import win32com.client  # type: ignore
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(Path(args.excel)))
        try:
            sh = wb.Worksheets("Orders")
            last = sh.Cells(sh.Rows.Count,1).End(-4162).Row  # xlUp
            print(last - 1)
        except Exception:
            print(0)
        finally:
            wb.Close(SaveChanges=False)
    finally:
        xl.Quit()


if __name__ == "__main__":
    main()

