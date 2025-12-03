#!/usr/bin/env python3
from __future__ import annotations

import argparse
import datetime as dt
import time
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Set

import jpholiday
import pandas as pd
from yahooquery import Ticker

DATA_ROOT = Path("data/raw/yahoo_1m")


def _default_batch_size(size: int) -> int:
    if size <= 50:
        return size
    if size <= 200:
        return 100
    return 200


def load_codes_from_csv(path: Path) -> List[str]:
    try:
        df = pd.read_csv(path)
    except Exception:
        return []
    for col in ("code", "Code", "ticker", "Ticker"):
        if col in df.columns:
            return df[col].dropna().astype(str).str.strip().tolist()
    return []


def gather_codes(args: argparse.Namespace) -> List[str]:
    codes: Set[str] = set()

    if args.codes:
        codes.update(c.strip() for c in args.codes if c.strip())

    for option in ("codes_file", "universe_file"):
        files = getattr(args, option, None)
        if not files:
            continue
        for file_path in files:
            path = Path(file_path)
            if path.exists():
                codes.update(load_codes_from_csv(path))

    if args.universe_glob:
        for matched in Path(".").glob(args.universe_glob):
            codes.update(load_codes_from_csv(matched))

    code_list = sorted({c for c in codes if c})
    if args.universe_limit and len(code_list) > args.universe_limit:
        code_list = code_list[: args.universe_limit]
    return code_list


def is_trading_day(day: dt.date) -> bool:
    return day.weekday() < 5 and not jpholiday.is_holiday(day)


def recent_trading_days(history_days: int, end_date: Optional[dt.date] = None) -> List[dt.date]:
    days: List[dt.date] = []
    current = end_date or dt.date.today()
    while len(days) < history_days:
        if is_trading_day(current):
            days.append(current)
        current -= dt.timedelta(days=1)
    return list(reversed(days))


def previous_trading_days(anchor: dt.date, count: int) -> List[dt.date]:
    days: List[dt.date] = []
    current = anchor - dt.timedelta(days=1)
    while len(days) < count:
        if is_trading_day(current):
            days.append(current)
        current -= dt.timedelta(days=1)
    return list(reversed(days))


def ensure_directory(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def normalize_history_frame(frame: pd.DataFrame) -> pd.DataFrame:
    df = frame.copy()
    if "symbol" in df.columns:
        df = df.rename(columns={"symbol": "code"})
    elif "Code" in df.columns:
        df = df.rename(columns={"Code": "code"})

    time_column = None
    for candidate in ("date", "datetime", "ts"):
        if candidate in df.columns:
            time_column = candidate
            break
    if time_column is None:
        return pd.DataFrame()

    df = df.rename(columns={time_column: "ts"})
    ts = pd.to_datetime(df["ts"], errors="coerce", utc=True)
    df["ts"] = ts.dt.tz_convert("Asia/Tokyo")
    return df


def save_day_history(df: pd.DataFrame, day: dt.date) -> int:
    if df.empty:
        return 0

    saved = 0
    if isinstance(df.index, pd.MultiIndex):
        for symbol, sub in df.groupby(level=0):
            local = normalize_history_frame(sub.droplevel(0).reset_index())
            if local.empty:
                continue
            local = local[local["ts"].dt.date == day]
            if local.empty:
                continue
            directory = DATA_ROOT / str(symbol)
            ensure_directory(directory)
            out_path = directory / f"{day:%Y-%m-%d}.parquet"
            try:
                local.set_index("ts").to_parquet(out_path)
                saved += 1
            except ImportError as exc:
                # parquet engine (pyarrow/fastparquet) not available; skip saving but do not abort the batch
                print(f"[update_minute_cache] parquet disabled for {symbol} {day}: {exc}")
            except Exception as exc:
                # any other failure should be logged but not crash the caller
                print(f"[update_minute_cache] failed to save parquet for {symbol} {day}: {exc}")
    else:
        frame = normalize_history_frame(df.reset_index())
        if "code" not in frame.columns:
            return 0
        for symbol, sub in frame.groupby("code"):
            local = sub[sub["ts"].dt.date == day]
            if local.empty:
                continue
            directory = DATA_ROOT / str(symbol)
            ensure_directory(directory)
            out_path = directory / f"{day:%Y-%m-%d}.parquet"
            try:
                local.set_index("ts").to_parquet(out_path)
                saved += 1
            except ImportError as exc:
                print(f"[update_minute_cache] parquet disabled for {symbol} {day}: {exc}")
            except Exception as exc:
                print(f"[update_minute_cache] failed to save parquet for {symbol} {day}: {exc}")
    return saved


def missing_codes_for_day(codes: Sequence[str], day: dt.date) -> List[str]:
    missing: List[str] = []
    for code in codes:
        path = DATA_ROOT / code / f"{day:%Y-%m-%d}.parquet"
        if not path.exists():
            missing.append(code)
    return missing


def chunked(sequence: Sequence[str], chunk_size: int) -> Iterable[List[str]]:
    for start in range(0, len(sequence), chunk_size):
        yield list(sequence[start : start + chunk_size])


def existing_days_for_code(code: str) -> List[dt.date]:
    directory = DATA_ROOT / code
    if not directory.exists():
        return []
    days: List[dt.date] = []
    for file in directory.glob("*.parquet"):
        try:
            day = dt.datetime.strptime(file.stem, "%Y-%m-%d").date()
            days.append(day)
        except ValueError:
            continue
    return sorted(days)


def fetch_day_for_codes(codes: Sequence[str], day: dt.date, batch_size: int, pause_seconds: float) -> int:
    if not codes:
        return 0

    fetched = 0
    start = day
    end = day + dt.timedelta(days=1)

    for chunk in chunked(codes, batch_size):
        try:
            history = Ticker(chunk, asynchronous=True).history(
                start=str(start), end=str(end), interval="1m"
            )
        except Exception:
            continue
        if isinstance(history, pd.DataFrame) and not history.empty:
            saved = save_day_history(history, day)
            fetched += saved
        if pause_seconds > 0:
            time.sleep(pause_seconds)
    return fetched


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--codes", nargs="*", help="Explicit ticker list", default=[])
    parser.add_argument("--codes-file", action="append", help="CSV with column 'code' or 'Ticker'")
    parser.add_argument("--universe-file", action="append", help="Additional CSV universe sources")
    parser.add_argument("--universe-glob", help="Glob pattern for CSV universes (e.g. data/universe/topvol_*.csv)")
    parser.add_argument("--universe-limit", type=int, default=0, help="Limit codes after gathering (0 = unlimited)")
    parser.add_argument("--history-days", type=int, default=5, help="Number of recent trading days to ensure")
    parser.add_argument("--start-date", help="Optional explicit start date (YYYY-MM-DD)")
    parser.add_argument("--end-date", help="Optional explicit end date (YYYY-MM-DD)")
    parser.add_argument("--batch-size", type=int, default=0, help="Ticker batch size per Yahoo request")
    parser.add_argument("--pause", type=float, default=0.5, help="Pause between batch requests in seconds")
    parser.add_argument(
        "--backfill-days",
        type=int,
        default=0,
        help="Number of older trading days to backfill (0 disables)",
    )
    args = parser.parse_args()

    codes = gather_codes(args)
    if not codes:
        print("update_minute_cache: no codes to process")
        return

    batch_size = args.batch_size if args.batch_size > 0 else _default_batch_size(len(codes))
    ensure_directory(DATA_ROOT)

    if args.start_date and args.end_date:
        start = dt.datetime.strptime(args.start_date, "%Y-%m-%d").date()
        end = dt.datetime.strptime(args.end_date, "%Y-%m-%d").date()
        recent_days: List[dt.date] = []
        current = start
        while current <= end:
            if is_trading_day(current):
                recent_days.append(current)
            current += dt.timedelta(days=1)
    else:
        recent_days = recent_trading_days(max(1, args.history_days))

    total_saved = 0
    for day in recent_days:
        missing = missing_codes_for_day(codes, day)
        if not missing:
            continue
        saved = fetch_day_for_codes(missing, day, batch_size, args.pause)
        total_saved += saved
        print(f"[update_minute_cache] {day}: saved {saved} tickers (missing {len(missing)})")

    if args.backfill_days > 0:
        backfill_map: Dict[dt.date, List[str]] = {}
        for code in codes:
            existing = existing_days_for_code(code)
            if not existing:
                continue
            oldest = existing[0]
            older_days = previous_trading_days(oldest, args.backfill_days)
            for day in older_days:
                path = DATA_ROOT / code / f"{day:%Y-%m-%d}.parquet"
                if path.exists():
                    continue
                backfill_map.setdefault(day, []).append(code)
        for day in sorted(backfill_map.keys()):
            saved = fetch_day_for_codes(backfill_map[day], day, batch_size, args.pause)
            total_saved += saved
            print(f"[update_minute_cache] backfill {day}: saved {saved} tickers (request {len(backfill_map[day])})")

    print(f"[update_minute_cache] completed for {len(codes)} codes, new files {total_saved}")


if __name__ == "__main__":
    main()
