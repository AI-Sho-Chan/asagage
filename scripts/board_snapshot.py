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


def read_named(wb, name: str):
    try:
        nr = wb.Names(name)
        ref = nr.RefersToRange
        return ref.Value
    except Exception:
        return None


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


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--board", default=r"excel/BoardLogger.xlsx")
    ap.add_argument("--outdir", default=r"output/board_logs")
    args = ap.parse_args()
    p = snapshot_board(Path(args.board), Path(args.outdir))
    print("written:", p)


if __name__ == "__main__":
    main()

