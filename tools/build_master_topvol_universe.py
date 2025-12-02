#!/usr/bin/env python3
from __future__ import annotations

import argparse
import datetime as dt
from pathlib import Path
from typing import Iterable, List, Set

import pandas as pd
import subprocess
import sys


HERE = Path(__file__).resolve().parent
REPO_ROOT = HERE.parent
UNIVERSE_DIR = REPO_ROOT / "data" / "universe"


def _default_source_paths() -> List[Path]:
    """Return default universe CSVs to union for the master universe.

    これらの CSV は日々の売買代金や値上がり/値下がりで絞り込んだ
    「流動性が高い銘柄」の集合として扱う。
    """
    candidates = [
        "dekidaka.csv",
        "dekidaka_kairi.csv",
        "dekidaka_kyuzo.csv",
        "neagari.csv",
        "nesagari.csv",
        "yori_neagari.csv",
        "yori_nesagari.csv",
        "tick.csv",
    ]
    paths: List[Path] = []
    for name in candidates:
        path = UNIVERSE_DIR / name
        if path.exists():
            paths.append(path)
    return paths


def _load_codes(path: Path) -> List[str]:
    try:
        df = pd.read_csv(path)
    except Exception:
        return []

    for col in ("code", "Code", "コード"):
        if col in df.columns:
            series = (
                df[col]
                .dropna()
                .astype(str)
                .str.strip()
                .str.upper()
            )
            codes = series.tolist()
            break
    else:
        return []

    out: List[str] = []
    for c in codes:
        if not c:
            continue
        if c.endswith(".T") or c.startswith("^"):
            out.append(c)
        else:
            out.append(f"{c}.T")
    return out


def build_master_codes(sources: Iterable[Path]) -> List[str]:
    seen: Set[str] = set()
    for src in sources:
        for code in _load_codes(src):
            if code:
                seen.add(code)
    return sorted(seen)


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Build a master weekly Top300 universe from all screened tickers.\n\n"
            "1) Union multiple daily universe CSVs under data/universe/.\n"
            "2) Use that union as --master for build_weekly_topvol_universe.py\n"
            "   to generate data/universe/topvol_<tag>.csv (weekly amount top N).\n"
        )
    )
    parser.add_argument(
        "--topn",
        type=int,
        default=300,
        help="How many codes to keep in weekly Top list (default 300)",
    )
    parser.add_argument(
        "--lookback",
        type=int,
        default=5,
        help="Trading days to sum for weekly amount (default 5)",
    )
    parser.add_argument(
        "--tag",
        default=dt.date.today().strftime("%Y%m%d"),
        help="Date tag for output filenames (default: today, YYYYMMDD)",
    )
    parser.add_argument(
        "--sources",
        nargs="*",
        help=(
            "Optional explicit list of CSVs to union for the master universe. "
            "If omitted, use standard screens in data/universe/*.csv "
            "(dekidaka, neagari, tick, etc.)."
        ),
    )
    args = parser.parse_args()

    if args.sources:
        source_paths = [
            Path(p) if Path(p).is_absolute() else (UNIVERSE_DIR / p) for p in args.sources
        ]
    else:
        source_paths = _default_source_paths()

    codes = build_master_codes(source_paths)
    if not codes:
        raise SystemExit("master_topvol_universe: no codes gathered from sources")

    master_dir = UNIVERSE_DIR
    master_dir.mkdir(parents=True, exist_ok=True)
    master_codes_path = master_dir / f"master_codes_{args.tag}.csv"

    df_master = pd.DataFrame({"code": codes})
    df_master.to_csv(master_codes_path, index=False)

    # Defer to existing weekly builder via subprocess to keep semantics consistent.
    cmd = [
        sys.executable,
        str((HERE / "build_weekly_topvol_universe.py").resolve()),
        "--master",
        str(master_codes_path),
        "--topn",
        str(args.topn),
        "--lookback",
        str(args.lookback),
        "--tag",
        args.tag,
    ]
    result = subprocess.run(cmd, cwd=str(REPO_ROOT))
    if result.returncode != 0:
        raise SystemExit(result.returncode)


if __name__ == "__main__":
    main()
