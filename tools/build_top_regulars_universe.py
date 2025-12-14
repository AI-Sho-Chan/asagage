#!/usr/bin/env python3
"""
「Top300（週間売買代金 上位）」のCSV（`data/universe/topvol_YYYYMMDD.csv`）が複数日ぶんある前提で、
「頻繁に上位に出てくる銘柄（Top200常連など）」を機械的に選んでCSVにします。

目的（かんたんに）:
- Yahooの1分足は遡れる日数が限られるため、消える前にローカルへ保存しておきたい
- その対象を「毎回よく出てくる銘柄」に絞ることで、保存作業を現実的にする

入力:
- `data/universe/topvol_*.csv`（列: `code`）

出力:
- `data/universe/top_regulars_<tag>.csv`（列: `code`）
- `data/universe/top_regulars_latest.csv`（直近のコピー）
- `data/universe/top_regulars_<tag>_stats.csv`（出現回数・平均順位のメモ）
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
    return ap.parse_args()


def main() -> None:
    args = parse_args()
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    files = _iter_recent_files(args.glob, args.lookback_files)
    df_stats = build_top_regulars(files)
    df_codes = df_stats.head(args.topn)[["code"]].copy()

    out_file = out_dir / f"top_regulars_{args.tag}.csv"
    df_codes.to_csv(out_file, index=False)

    latest = out_dir / "top_regulars_latest.csv"
    shutil.copyfile(out_file, latest)

    stats_file = out_dir / f"top_regulars_{args.tag}_stats.csv"
    df_stats.head(args.topn).to_csv(stats_file, index=False)

    print(
        f"[top_regulars] files={len(files)} topn={len(df_codes)} out={out_file} latest={latest} stats={stats_file}"
    )


if __name__ == "__main__":
    main()

