#!/usr/bin/env python3
"""
Build a small "always-cache" universe from recent weekly TopVol CSVs.

What it does
------------
- Reads multiple `data/universe/topvol_YYYYMMDD.csv` files (Top300 by weekly turnover).
- Counts how often each ticker appears and its average rank.
- Outputs the "Top N regulars" (default Top200) as:
  - `data/universe/top_regulars_<tag>.csv`
  - `data/universe/top_regulars_latest.csv` (copy)
  - `data/universe/top_regulars_<tag>_stats.csv`

Why we want this
----------------
Yahoo 1-minute data is limited in how far back we can fetch.
By continuously saving 1-minute bars for frequent TopVol tickers, the weekend batch:
- has fewer missing days
- downloads less data
- runs more stably and faster
"""

from __future__ import annotations

import argparse
import datetime as dt
import re
import shutil
from pathlib import Path
from typing import Dict, Iterable, List, Tuple

import pandas as pd

DEFAULT_GLOB = "data/universe/topvol_*.csv"
DEFAULT_OUT_DIR = Path("data/universe")


def _parse_date_tag(name: str) -> str | None:
    match = re.search(r"topvol_(\d{8})\.csv$", name, flags=re.IGNORECASE)
    return match.group(1) if match else None


def _load_codes(path: Path) -> List[str]:
    try:
        df = pd.read_csv(path)
    except Exception:
        return []
    if "code" not in df.columns:
        return []
    return df["code"].dropna().astype(str).str.strip().str.upper().tolist()


def _iter_recent_files(glob_pattern: str, lookback_files: int) -> List[Path]:
    files: List[Tuple[str, Path]] = []
    for path in Path(".").glob(glob_pattern):
        if path.is_dir():
            continue
        tag = _parse_date_tag(path.name)
        if not tag:
            continue
        if "TEST" in path.name.upper():
            continue
        files.append((tag, path))
    files.sort(key=lambda item: item[0])
    if lookback_files > 0:
        files = files[-lookback_files:]
    return [path for _, path in files]


def build_top_regulars(files: Iterable[Path]) -> pd.DataFrame:
    counts: Dict[str, int] = {}
    rank_sum: Dict[str, int] = {}
    seen_files = 0

    for path in files:
        codes = _load_codes(path)
        if not codes:
            continue
        seen_files += 1
        for idx, code in enumerate(codes, start=1):
            counts[code] = counts.get(code, 0) + 1
            rank_sum[code] = rank_sum.get(code, 0) + idx

    if seen_files == 0:
        return pd.DataFrame(columns=["code", "appear_count", "avg_rank"])

    rows = []
    for code, count in counts.items():
        avg_rank = rank_sum.get(code, 0) / max(1, count)
        rows.append((code, count, avg_rank))

    df = pd.DataFrame(rows, columns=["code", "appear_count", "avg_rank"])
    return df.sort_values(
        ["appear_count", "avg_rank", "code"],
        ascending=[False, True, True],
    ).reset_index(drop=True)


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser()
    ap.add_argument("--glob", default=DEFAULT_GLOB, help=f"Input glob (default: {DEFAULT_GLOB})")
    ap.add_argument("--lookback-files", type=int, default=20, help="How many topvol files to use (default 20)")
    ap.add_argument("--topn", type=int, default=200, help="How many codes to output (default 200)")
    ap.add_argument(
        "--tag",
        default=dt.date.today().strftime("%Y%m%d"),
        help="Date tag for output filename (default: today, YYYYMMDD)",
    )
    ap.add_argument("--out-dir", default=str(DEFAULT_OUT_DIR), help="Output directory (default: data/universe)")
    ap.add_argument(
        "--update-ever",
        action="store_true",
        help=(
            "Also update a persistent universe file that keeps any ticker that has ever appeared "
            "in the TopN regulars. This grows over time but avoids losing cached history when a ticker "
            "temporarily drops out of the TopN."
        ),
    )
    ap.add_argument(
        "--ever-file",
        default="top_regulars_ever.csv",
        help="Filename (within out-dir) for the persistent universe (default: top_regulars_ever.csv)",
    )
    return ap.parse_args()


def main() -> None:
    args = parse_args()
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    files = _iter_recent_files(args.glob, args.lookback_files)
    df_stats = build_top_regulars(files)
    df_codes = df_stats.head(args.topn)[["code"]].copy()

    out_file = out_dir / f"top_regulars_{args.tag}.csv"
    df_codes.to_csv(out_file, index=False, encoding="utf-8-sig")

    latest = out_dir / "top_regulars_latest.csv"
    shutil.copyfile(out_file, latest)

    stats_file = out_dir / f"top_regulars_{args.tag}_stats.csv"
    df_stats.head(args.topn).to_csv(stats_file, index=False, encoding="utf-8-sig")

    ever_updated = ""
    if args.update_ever:
        ever_path = out_dir / str(args.ever_file)
        existing: List[str] = []
        if ever_path.exists():
            try:
                prev = pd.read_csv(ever_path)
                if "code" in prev.columns:
                    existing = prev["code"].dropna().astype(str).str.strip().str.upper().tolist()
            except Exception:
                existing = []
        merged = sorted(set(existing).union(df_codes["code"].dropna().astype(str).str.strip().str.upper().tolist()))
        pd.DataFrame({"code": merged}).to_csv(ever_path, index=False, encoding="utf-8-sig")
        added = max(0, len(merged) - len(set(existing)))
        ever_updated = f" ever={ever_path} (total={len(merged)} added={added})"

    print(
        f"[top_regulars] files={len(files)} topn={len(df_codes)} out={out_file} latest={latest} stats={stats_file}{ever_updated}"
    )


if __name__ == "__main__":
    main()
