from __future__ import annotations

import argparse
import datetime as dt
from pathlib import Path
from typing import List, Optional

import math
import pandas as pd


def wilson_ci(k: float, n: float, z: float = 1.96) -> tuple[float, float]:
    if n <= 0:
        return (0.0, 0.0)
    p = k / n
    denom = 1.0 + (z * z) / n
    center = (p + (z * z) / (2 * n)) / denom
    half = (z * math.sqrt((p * (1 - p) / n) + (z * z) / (4 * n * n))) / denom
    lo = max(0.0, center - half)
    hi = min(1.0, center + half)
    return (lo, hi)


def find_latest_nightly(root: Path) -> Optional[Path]:
    cands = sorted([p for p in root.glob('NIGHTLY_*') if p.is_dir()])
    return cands[-1] if cands else None


def load_grids(run_root: Path) -> pd.DataFrame:
    files: List[Path] = list(run_root.rglob("_GRID_FULL.csv"))
    frames: List[pd.DataFrame] = []
    for f in files:
        try:
            df = pd.read_csv(f)
            df["_src"] = str(f)
            frames.append(df)
        except Exception:
            continue
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def ensure_numeric(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')
    return df


def group_by_j(df: pd.DataFrame) -> pd.DataFrame:
    # pick columns if present
    df = ensure_numeric(df, [
        'forward_trades','forward_winrate','forward_pf_eff','forward_exp_boot_mean'
    ])
    if 'J_th' not in df.columns:
        return pd.DataFrame()
    g = df.groupby('J_th', as_index=False).agg(
        grids=('J_th','count'),
        trades=('forward_trades','sum'),
        win_sum=('forward_winrate','sum'),
        win_mean=('forward_winrate','mean'),
        pf_mean=('forward_pf_eff','mean'),
        exp_mean=('forward_exp_boot_mean','mean'),
    )
    # Wilson CI uses wins/trades; approximate wins = winrate * trades aggregated per grid
    # This is an approximation; better is raw wins per trade, but not available across files.
    # We fallback to using win_mean and trades as scale for indicative CI only.
    wins_approx = []
    ci_lo = []
    ci_hi = []
    for _, row in g.iterrows():
        tr = float(row['trades']) if pd.notna(row['trades']) else 0.0
        wr = float(row['win_mean']) if pd.notna(row['win_mean']) else 0.0
        k = max(0.0, min(tr, wr * tr))
        lo, hi = wilson_ci(k, tr) if tr > 0 else (0.0, 0.0)
        wins_approx.append(k)
        ci_lo.append(lo)
        ci_hi.append(hi)
    g['wins_approx'] = wins_approx
    g['win_ci_low'] = ci_lo
    g['win_ci_high'] = ci_hi
    return g.sort_values('J_th')


def pivot_2d(df: pd.DataFrame, idx: str, col: str, val: str, agg: str='mean') -> pd.DataFrame:
    if idx not in df.columns or col not in df.columns or val not in df.columns:
        return pd.DataFrame()
    df = ensure_numeric(df, [val])
    pt = pd.pivot_table(df, index=idx, columns=col, values=val, aggfunc=agg)
    return pt


def write_outputs(out_dir: Path, byj: pd.DataFrame, df: pd.DataFrame) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    if not byj.empty:
        byj.to_csv(out_dir / 'by_J_th.csv', index=False, encoding='utf-8-sig')
    # Heatmaps as CSVs
    for v in ['forward_winrate','forward_pf_eff','forward_exp_boot_mean']:
        for second in ['ATR_n','TPk','SLk']:
            pt = pivot_2d(df, 'J_th', second, v, 'mean')
            if not pt.empty:
                pt.to_csv(out_dir / f'heat_{v}_by_J_th_x_{second}.csv', encoding='utf-8-sig')


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument('--root', help='Path to NIGHTLY_YYYYMMDD folder (default: latest under output/bt30)')
    ap.add_argument('--out', help='Reports output folder (default: reports/param_stats/NIGHTLY_xxx)')
    args = ap.parse_args()

    base = Path('output/bt30')
    run_root = Path(args.root) if args.root else find_latest_nightly(base)
    if not run_root or not run_root.exists():
        raise SystemExit('No NIGHTLY folder found')
    df = load_grids(run_root)
    if df.empty:
        raise SystemExit('No _GRID_FULL.csv found to analyze')

    byj = group_by_j(df)
    date_tag = run_root.name.replace('NIGHTLY_','')
    out_dir = Path(args.out) if args.out else Path('reports/param_stats') / run_root.name
    write_outputs(out_dir, byj, df)

    summary = {
        'run_root': str(run_root.resolve()),
        'grids': int(len(df.index)),
        'j_bins': int(len(byj.index) if not byj.empty else 0),
        'out_dir': str(out_dir.resolve()),
        'generated_at': dt.datetime.now().isoformat(),
    }
    pd.Series(summary).to_json(out_dir / 'summary.json', force_ascii=False, indent=2)
    print('Analysis written to', out_dir)


if __name__ == '__main__':
    main()

