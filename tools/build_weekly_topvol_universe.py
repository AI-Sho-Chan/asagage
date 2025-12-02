#!/usr/bin/env python3
"""
週末バッチ用: 直近5営業日の「終値×出来高」合計で上位N銘柄を作る。

できるだけシンプルに、日足データで週間売買代金トップを算出し、
data/universe/topvol_<date>.csv (code 列のみ) を生成する。

想定の呼び出し先:
- GCP VM の run_weekend_only.sh 冒頭で実行し、その週の Top300 を常に更新。
"""

from __future__ import annotations

import argparse
import datetime as dt
from pathlib import Path
from typing import Iterable, List, Optional

import jpholiday
import pandas as pd
from yahooquery import Ticker

DEFAULT_OUT_DIR = Path("data/universe")
DEFAULT_MASTER = Path("data/universe/tick.csv")


def is_trading_day(day: dt.date) -> bool:
    return day.weekday() < 5 and not jpholiday.is_holiday(day)


def recent_trading_days(count: int, end: Optional[dt.date] = None) -> List[dt.date]:
    days: List[dt.date] = []
    current = end or dt.date.today()
    while len(days) < count:
        if is_trading_day(current):
            days.append(current)
        current -= dt.timedelta(days=1)
    return list(reversed(days))


def iter_chunks(seq: List[str], size: int) -> Iterable[List[str]]:
    for pos in range(0, len(seq), size):
        yield seq[pos : pos + size]


def load_master_codes(path: Path) -> List[str]:
    if not path.exists():
        return []
    try:
        df = pd.read_csv(path)
    except Exception:
        return []
    for col in ("code", "Code", "ticker", "Ticker"):
        if col in df.columns:
            codes = (
                df[col]
                .dropna()
                .astype(str)
                .str.strip()
                .str.upper()
                .tolist()
            )
            return [c if c.endswith(".T") or c.startswith("^") else f"{c}.T" for c in codes]
    return []


def fetch_daily_ohlc(codes: List[str], start: dt.date, end: dt.date, batch: int) -> pd.DataFrame:
    frames: List[pd.DataFrame] = []
    for chunk in iter_chunks(codes, batch):
        try:
            hist = Ticker(chunk, asynchronous=True).history(
                start=str(start), end=str(end), interval="1d"
            )
        except Exception:
            continue
        if isinstance(hist, pd.DataFrame) and not hist.empty:
            frames.append(hist)
    if not frames:
        return pd.DataFrame()
    df = pd.concat(frames, ignore_index=False)
    if not isinstance(df.index, pd.MultiIndex):
        return pd.DataFrame()
    df = df.reset_index().rename(columns={"symbol": "code"})
    df["code"] = df["code"].astype(str).str.upper()
    # Yahoo 側の仕様変更で tz 付き / なしが混在するケースがあるため、utc=True で正規化する
    df["date"] = pd.to_datetime(df["date"], utc=True).dt.date
    keep_cols = {"code", "date", "close", "volume"}
    df = df[[c for c in df.columns if c in keep_cols]]
    return df


def build_weekly_topvol(args: argparse.Namespace) -> Path:
    master = Path(args.master)
    codes = load_master_codes(master)
    if not codes:
        raise SystemExit(f"master code list not found or empty: {master}")

    anchor_date = None
    if args.anchor:
        anchor_date = dt.datetime.strptime(args.anchor, "%Y-%m-%d").date()
    days = recent_trading_days(args.lookback, end=anchor_date)
    start = days[0]
    end = days[-1] + dt.timedelta(days=1)

    daily = fetch_daily_ohlc(codes, start, end, args.batch_size)
    if daily.empty:
        raise SystemExit("failed to fetch daily data for master universe")

    daily["amt"] = pd.to_numeric(daily["close"], errors="coerce") * pd.to_numeric(
        daily["volume"], errors="coerce"
    )
    weekly = (
        daily.groupby("code")["amt"]
        .sum()
        .reset_index(name="amt_sum")
        .sort_values("amt_sum", ascending=False)
    )
    topn = weekly.head(args.topn).copy()
    topn = topn[["code"]]

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_file = out_dir / f"topvol_{args.tag}.csv"
    topn.to_csv(out_file, index=False)
    print(f"[build_weekly_topvol] saved {len(topn)} codes to {out_file}")
    return out_file


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--master", default=DEFAULT_MASTER, help="CSV with master code list (default: data/universe/tick.csv)")
    parser.add_argument("--topn", type=int, default=300, help="How many codes to keep (default 300)")
    parser.add_argument("--lookback", type=int, default=5, help="Trading days to sum (default 5)")
    parser.add_argument("--out-dir", default=DEFAULT_OUT_DIR, help="Directory to write topvol CSV")
    parser.add_argument("--tag", default=dt.date.today().strftime("%Y%m%d"), help="Date tag for output filename")
    parser.add_argument("--batch-size", type=int, default=80, help="YahooQuery batch size (default 80)")
    parser.add_argument("--anchor", help="Optional anchor date YYYY-MM-DD (default: today)")
    return parser.parse_args()


if __name__ == "__main__":
    build_weekly_topvol(parse_args())
