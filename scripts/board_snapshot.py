import argparse
import csv
import datetime as dt
from pathlib import Path
from typing import Dict, List

import win32com.client  # type: ignore


NAMES = [
    "Ticker",
    *[f"BID_P_{i}" for i in range(10)],
    *[f"BID_Q_{i}" for i in range(10)],
    *[f"ASK_P_{i}" for i in range(10)],
    *[f"ASK_Q_{i}" for i in range(10)],
    "TOP3_AMT",
    "TOP10_AMT",
]

DASH_HEADERS = [
    "Ticker",
    "J",
    "J_th",
    "J_ratio",
    "EntrySide",
    "Selected",
    "session",
    "NKY_allowed_side",
    "driver_allowed_side",
    "driver_day_trend",
    "driver_window_trend",
    "trend_driver",
    "trend_window",
    "trend_bp_th",
    "trend_allowed_policy",
]

XL_TO_LEFT = -4159
XL_UP = -4162


def read_named(wb, name: str):
    try:
        nr = wb.Names(name)
        ref = nr.RefersToRange
        return ref.Value
    except Exception:
        return None


def _build_header_map(ws, header_row: int) -> Dict[str, int]:
    last_col = ws.Cells(header_row, ws.Columns.Count).End(XL_TO_LEFT).Column
    mapping: Dict[str, int] = {}
    for col in range(1, last_col + 1):
        val = str(ws.Cells(header_row, col).Value or "").strip()
        if val:
            mapping[val] = col
    return mapping


def snapshot_board(xlsx_path: Path, outdir: Path) -> Path:
    outdir.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(xlsx_path.resolve()))
        row: Dict[str, object] = {"ts": dt.datetime.now().isoformat()}
        for nm in NAMES:
            row[nm] = read_named(wb, nm)
        ticker = str(row.get("Ticker") or "NA").strip()
        ts_tag = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        out = outdir / f"{ticker}_{ts_tag}_book.csv"
        with out.open("w", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=list(row.keys()))
            w.writeheader()
            w.writerow(row)
        return out
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
        excel.Quit()


def snapshot_dashboard_j(xls_path: Path, outdir: Path) -> Path | None:
    outdir.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(xls_path.resolve()))
        try:
            ws = wb.Worksheets("NewDashboardV2")
        except Exception:
            return None

        header_map = _build_header_map(ws, 5)
        cols = {name: header_map.get(name) for name in DASH_HEADERS}
        if cols["Ticker"] is None:
            return None

        last_row = ws.Cells(ws.Rows.Count, cols["Ticker"]).End(XL_UP).Row
        rows: List[Dict[str, object]] = []
        for row in range(6, last_row + 1):
            ticker = str(ws.Cells(row, cols["Ticker"]).Value or "").strip()
            if not ticker:
                continue
            record: Dict[str, object] = {"ts": dt.datetime.now().isoformat()}
            for name, col in cols.items():
                if col is None:
                    record[name] = ""
                else:
                    record[name] = ws.Cells(row, col).Value
            rows.append(record)

        if not rows:
            return None

        ts_tag = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        out = outdir / f"dashboard_j_{ts_tag}.csv"
        with out.open("w", newline="", encoding="utf-8-sig") as fh:
            writer = csv.DictWriter(fh, fieldnames=["ts", *DASH_HEADERS])
            writer.writeheader()
            writer.writerows(rows)
        return out
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
        excel.Quit()


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--board", default=r"excel/BoardLogger.xlsx")
    ap.add_argument("--outdir", default=r"output/board_logs")
    args = ap.parse_args()
    p = snapshot_board(Path(args.board), Path(args.outdir))
    print("written:", p)


if __name__ == "__main__":
    main()
