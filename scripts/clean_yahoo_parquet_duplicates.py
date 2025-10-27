from __future__ import annotations

import argparse
from pathlib import Path
from typing import Dict, Iterable

import pandas as pd


NEEDED_ORDER = ["open", "high", "low", "close", "volume"]


def normalize_frame(df: pd.DataFrame) -> pd.DataFrame:
    """Return a DataFrame containing only OHLCV columns with unique labels."""
    if isinstance(df.columns, pd.MultiIndex):
        selected: Dict[str, pd.Series] = {}
        for tup in df.columns:
            name = str(tup[0]).strip().lower()
            if name == "adj close":
                continue
            if name in NEEDED_ORDER and name not in selected:
                selected[name] = df[tup]
        df = pd.DataFrame(selected, index=df.index)
    else:
        cols = [str(c).strip().lower() for c in df.columns]
        df.columns = cols
        selected = {}
        for name in cols:
            if name == "adj close":
                continue
            if name in NEEDED_ORDER and name not in selected:
                selected[name] = df[name]
        df = pd.DataFrame(selected, index=df.index)

    if not set(NEEDED_ORDER).issubset(df.columns):
        raise ValueError("Missing required OHLCV columns after normalization")

    df = df[NEEDED_ORDER].copy()
    df.index = pd.to_datetime(df.index)
    return df


def iter_parquet_files(root: Path) -> Iterable[Path]:
    for ticker_dir in root.iterdir():
        if not ticker_dir.is_dir():
            continue
        yield from ticker_dir.glob("*.parquet")


def main() -> None:
    ap = argparse.ArgumentParser(description="Normalize Yahoo 1m parquet files (dedupe columns).")
    ap.add_argument(
        "--root",
        default="data/raw/yahoo_1m",
        help="Root directory containing <ticker>/<date>.parquet files",
    )
    ap.add_argument(
        "--dry-run",
        action="store_true",
        help="Scan only and report files requiring fixes without rewriting.",
    )
    args = ap.parse_args()

    root = Path(args.root).resolve()
    if not root.exists():
        print(f"Root {root} does not exist.")
        return

    fixed = 0
    skipped = 0
    for path in iter_parquet_files(root):
        try:
            df = pd.read_parquet(path)
        except Exception as exc:  # pragma: no cover
            print(f"[SKIP] {path}: failed to read ({exc})")
            skipped += 1
            continue

        try:
            normalized = normalize_frame(df)
        except ValueError:
            skipped += 1
            continue

        # If already normalized and columns unique, skip rewrite.
        if set(normalized.columns) == set(df.columns if not isinstance(df.columns, pd.MultiIndex) else []):
            continue

        fixed += 1
        if args.dry_run:
            print(f"[DRY] would normalize {path}")
        else:
            normalized.to_parquet(path, index=True)
            print(f"[FIXED] {path}")

    print(f"Completed. fixed={fixed}, skipped={skipped}")


if __name__ == "__main__":
    main()

