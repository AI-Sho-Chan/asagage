import argparse
import datetime as dt
import time
from pathlib import Path
from typing import Iterable, List, Set, Tuple

import pandas as pd
import yfinance as yf


def collect_codes(universe_dir: Path, extra_csv: Iterable[Path]) -> List[str]:
    codes: Set[str] = set()
    # universe CSVs
    for csv_path in universe_dir.glob("*.csv"):
        try:
            df = pd.read_csv(csv_path)
        except Exception:
            continue
        for col in ("code", "Code", "ticker", "Ticker"):
            if col in df.columns:
                vals = df[col].dropna().astype(str)
                for v in vals:
                    v = v.strip()
                    if not v:
                        continue
                    if not v.endswith(".T") and v.isdigit() and len(v) == 4:
                        v = f"{v}.T"
                    codes.add(v)
                break

    for csv_path in extra_csv:
        if not csv_path.exists():
            continue
        try:
            df = pd.read_csv(csv_path)
        except Exception:
            continue
        for col in ("Ticker", "ticker", "code", "Code"):
            if col in df.columns:
                vals = df[col].dropna().astype(str)
                for v in vals:
                    v = v.strip()
                    if not v:
                        continue
                    if not v.endswith(".T") and v.isdigit() and len(v) == 4:
                        v = f"{v}.T"
                    codes.add(v)
                break
    return sorted(codes)


def fetch_ticker(ticker: str, days: int, sleep: float, raw_root: Path) -> Tuple[int, str]:
    end = dt.datetime.now(dt.timezone.utc) + dt.timedelta(days=1)
    start = end - dt.timedelta(days=days)
    earliest = end - dt.timedelta(days=30)
    if start < earliest:
        start = earliest
    chunk = dt.timedelta(days=7)
    saved = 0
    out_dir = raw_root / ticker
    out_dir.mkdir(parents=True, exist_ok=True)
    first_chunk = True
    missing_reason = ""

    existing_dates = [p.stem for p in out_dir.glob("*.parquet") if p.is_file()]
    if existing_dates:
        try:
            latest = max(existing_dates)
            latest_date = dt.date.fromisoformat(latest)
            next_day = latest_date + dt.timedelta(days=1)
            start_candidate = dt.datetime.combine(next_day, dt.time(0), tzinfo=dt.timezone.utc)
            if start_candidate > start:
                start = start_candidate
        except Exception:
            pass

    cur = start
    while cur < end:
        chunk_end = min(cur + chunk, end)
        try:
            df = yf.download(
                ticker,
                start=cur,
                end=chunk_end,
                interval="1m",
                auto_adjust=False,
                prepost=False,
                progress=False,
            )
        except Exception as exc:
            missing_reason = str(exc)
            df = pd.DataFrame()
        if not df.empty:
            try:
                if df.index.tz is None:
                    df.index = df.index.tz_localize("UTC").tz_convert("Asia/Tokyo")
                else:
                    df.index = df.index.tz_convert("Asia/Tokyo")
            except Exception:
                pass
            df = df[~df.index.duplicated(keep="last")]
            df["date"] = df.index.date
            for date, g in df.groupby("date"):
                fname = out_dir / f"{date}.parquet"
                if fname.exists():
                    continue
            g.drop(columns=["date"], errors="ignore").to_parquet(fname)
            saved += len(g)
        time.sleep(sleep)
        cur = chunk_end
        if first_chunk and saved == 0 and df.empty:
            if not missing_reason:
                missing_reason = "no price data"
            break
        first_chunk = False
    return saved, missing_reason


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--days", type=int, default=90)
    ap.add_argument("--universe", default="data/universe")
    ap.add_argument("--extra", action="append", default=["output/excel/candidates_nextday.csv"])
    ap.add_argument("--raw-root", default="data/raw/yahoo_1m")
    ap.add_argument("--sleep", type=float, default=0.4)
    args = ap.parse_args()

    universe_dir = Path(args.universe)
    extra_csv = [Path(p) for p in args.extra]
    raw_root = Path(args.raw_root)

    tickers = collect_codes(universe_dir, extra_csv)
    if not tickers:
        raise SystemExit("No tickers found to update")

    print(f"Updating {len(tickers)} tickers (days={args.days})")
    missing_log: List[str] = []
    for i, ticker in enumerate(tickers, 1):
        saved = 0
        reason = ""
        try:
            saved, reason = fetch_ticker(ticker, args.days, args.sleep, raw_root)
        except Exception as exc:  # pragma: no cover
            reason = str(exc)
        if saved == 0 and reason:
            msg = f"[{i}/{len(tickers)}] {ticker}: no data ({reason})"
            missing_log.append(msg)
        else:
            msg = f"[{i}/{len(tickers)}] {ticker}: saved {saved} rows"
        print(msg)

    if missing_log:
        log_path = raw_root / "_missing.log"
        with log_path.open("a", encoding="utf-8") as fh:
            fh.write(f"# {dt.datetime.now():%Y-%m-%d %H:%M:%S}\n")
            for line in missing_log:
                fh.write(line + "\n")


if __name__ == "__main__":
    main()
