#!/usr/bin/env python3
from __future__ import annotations

import argparse
import datetime as dt
from pathlib import Path
import sys

import jpholiday

DATA_ROOT = Path("data/raw/yahoo_1m")
REFERENCE_TICKERS = ["1301.T", "1332.T", "7203.T", "9984.T"]
REMOTE_LOOKBACK_DAYS = 45  # Yahoo 1m history limit is ~60 days


def has_local_data(target: dt.date) -> bool:
    date_name = f"{target:%Y-%m-%d}.parquet"
    for ticker in REFERENCE_TICKERS:
        candidate = DATA_ROOT / ticker / date_name
        if candidate.exists():
            return True
    return False


def fetch_remote_minutes(target: dt.date) -> bool:
    """Attempt to download 1m data for reference tickers and persist locally.

    Returns True when at least one ticker produced data for the target date.
    """
    # Yahoo only retains ~60 days of minute data; avoid unnecessary calls.
    if (dt.date.today() - target).days > REMOTE_LOOKBACK_DAYS:
        return False

    try:
        from yahooquery import Ticker  # type: ignore
        import pandas as pd  # type: ignore
    except Exception:
        return False

    start = target
    end = target + dt.timedelta(days=1)
    try:
        hist = Ticker(REFERENCE_TICKERS, asynchronous=True).history(
            start=str(start), end=str(end), interval="1m"
        )
    except Exception:
        return False

    if not isinstance(hist, pd.DataFrame) or hist.empty:
        return False

    saved = False

    def normalize_frame(frame: pd.DataFrame) -> pd.DataFrame:
        local = frame.copy()
        ts_col = None
        for cand in ("date", "datetime", "ts"):
            if cand in local.columns:
                ts_col = cand
                break
        if ts_col is None:
            return pd.DataFrame()
        local = local.rename(columns={ts_col: "ts"})
        ts = pd.to_datetime(local["ts"], errors="coerce", utc=True)
        local["ts"] = ts.dt.tz_convert("Asia/Tokyo")
        return local

    try:
        if isinstance(hist.index, pd.MultiIndex):
            for symbol, sub in hist.groupby(level=0):
                frame = sub.droplevel(0).reset_index()
                df = normalize_frame(frame)
                if df.empty:
                    continue
                df["ts"] = pd.to_datetime(df["ts"])
                df = df[df["ts"].dt.date == target]
                if df.empty:
                    continue
                out_dir = DATA_ROOT / symbol
                out_dir.mkdir(parents=True, exist_ok=True)
                out_path = out_dir / f"{target:%Y-%m-%d}.parquet"
                df.set_index("ts").to_parquet(out_path)
                saved = True
        else:
            frame = hist.reset_index()
            symbol_col = None
            for cand in ("symbol", "code"):
                if cand in frame.columns:
                    symbol_col = cand
                    break
            if symbol_col is None:
                return False
            for symbol, sub in frame.groupby(symbol_col):
                df = normalize_frame(sub)
                if df.empty:
                    continue
                df["ts"] = pd.to_datetime(df["ts"])
                df = df[df["ts"].dt.date == target]
                if df.empty:
                    continue
                out_dir = DATA_ROOT / symbol
                out_dir.mkdir(parents=True, exist_ok=True)
                out_path = out_dir / f"{target:%Y-%m-%d}.parquet"
                df.set_index("ts").to_parquet(out_path)
                saved = True
    except Exception:
        return False

    return saved


def is_calendar_open(target: dt.date) -> bool:
    if target.weekday() >= 5:
        return False
    if jpholiday.is_holiday(target):
        return False
    if target.month == 1 and target.day in (1, 2, 3):
        return False
    if target.month == 12 and target.day == 31:
        return False
    return True


def is_trading_day(target: dt.date) -> bool:
    if not is_calendar_open(target):
        return False
    if has_local_data(target):
        return True
    return fetch_remote_minutes(target)


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", required=True, help="ISO date (YYYY-MM-DD)")
    args = parser.parse_args(argv)

    try:
        target = dt.datetime.strptime(args.date, "%Y-%m-%d").date()
    except ValueError:
        print("invalid date format", file=sys.stderr)
        return 2

    if is_trading_day(target):
        print("trading-day")
        return 0

    print("holiday", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
