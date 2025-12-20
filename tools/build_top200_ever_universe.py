#!/usr/bin/env python3
"""
Maintain a persistent "ever appeared in weekly TopN" universe.

Requested policy:
  - If a ticker ever appears in the weekly Top200 list, keep it in the
    persistent list forever (so we keep accumulating 1-minute data for it).

This script updates:
  - data/universe/top200_ever.csv (default, persistent)
  - data/universe/top200_ever_latest.csv (copy of persistent, convenient alias)
  - data/universe/top200_ever_<tag>.csv (optional dated snapshot)

It works incrementally:
  merged = union(existing_ever, latest_week_topN)
"""

from __future__ import annotations

import argparse
import datetime as dt
import re
import shutil
from pathlib import Path
from typing import List, Tuple

import pandas as pd

DEFAULT_GLOB = "data/universe/topvol_*.csv"
DEFAULT_OUT_DIR = Path("data/universe")


def _parse_date_tag(name: str) -> str | None:
    match = re.search(r"topvol_(\d{8})\.csv$", name, flags=re.IGNORECASE)
    return match.group(1) if match else None


def _pick_latest_topvol(glob_pattern: str) -> Tuple[str, Path] | None:
    items: List[Tuple[str, Path]] = []
    for path in Path(".").glob(glob_pattern):
        if path.is_dir():
            continue
        tag = _parse_date_tag(path.name)
        if not tag:
            continue
        if "TEST" in path.name.upper():
            continue
        items.append((tag, path))
    if not items:
        return None
    items.sort(key=lambda it: it[0])
    return items[-1]


def _load_codes(path: Path) -> List[str]:
    try:
        df = pd.read_csv(path)
    except Exception:
        return []
    if "code" not in df.columns:
        return []
    return df["code"].dropna().astype(str).str.strip().str.upper().tolist()


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser()
    ap.add_argument("--glob", default=DEFAULT_GLOB, help=f"Input glob (default: {DEFAULT_GLOB})")
    ap.add_argument("--topn", type=int, default=200, help="Weekly TopN to union into ever list (default 200)")
    ap.add_argument(
        "--tag",
        default=dt.date.today().strftime("%Y%m%d"),
        help="Date tag for optional snapshot filename (YYYYMMDD). Default: today.",
    )
    ap.add_argument("--out-dir", default=str(DEFAULT_OUT_DIR), help="Output directory (default: data/universe)")
    ap.add_argument("--ever-file", default="top200_ever.csv", help="Persistent universe filename (default: top200_ever.csv)")
    ap.add_argument("--write-snapshot", action="store_true", help="Also write a dated snapshot file top200_ever_<tag>.csv")
    return ap.parse_args()


def main() -> None:
    args = parse_args()
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    picked = _pick_latest_topvol(args.glob)
    if not picked:
        raise SystemExit(f"No topvol files found for glob: {args.glob}")

    src_tag, src_path = picked
    week_codes = _load_codes(src_path)[: int(args.topn)]
    if not week_codes:
        raise SystemExit(f"Empty topvol codes: {src_path}")

    ever_path = out_dir / str(args.ever_file)
    existing: List[str] = _load_codes(ever_path) if ever_path.exists() else []

    merged = sorted(set(existing).union(set(week_codes)))
    pd.DataFrame({"code": merged}).to_csv(ever_path, index=False, encoding="utf-8-sig")

    latest = out_dir / "top200_ever_latest.csv"
    shutil.copyfile(ever_path, latest)

    if args.write_snapshot:
        snap = out_dir / f"top200_ever_{args.tag}.csv"
        shutil.copyfile(ever_path, snap)

    print(
        {
            "source_topvol": str(src_path),
            "source_tag": src_tag,
            "topn_week": int(args.topn),
            "week_codes": len(week_codes),
            "ever_total": len(merged),
            "ever_added": max(0, len(merged) - len(set(existing))),
            "out_ever": str(ever_path),
            "out_latest": str(latest),
        }
    )


if __name__ == "__main__":
    main()

