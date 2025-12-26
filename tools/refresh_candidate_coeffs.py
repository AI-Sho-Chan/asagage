#!/usr/bin/env python3
"""Refresh per-ticker coeffs/corrs for an existing candidates CSV.

This avoids rerunning the heavy weekend WF pipeline when only
BiasSlope/CorrSlope/CorrNKY/CorrTOPIX need to be anchored to a date.
"""
from __future__ import annotations

import argparse
import datetime as dt
import importlib.util
import shutil
import subprocess
import sys
from pathlib import Path

import pandas as pd

def load_nightly_helpers(repo_root: Path):
    module_path = repo_root / "scripts" / "nightly_build_candidates.py"
    if not module_path.exists():
        raise SystemExit(f"missing nightly_build_candidates.py at {module_path}")
    spec = importlib.util.spec_from_file_location("nightly_build_candidates", module_path)
    if spec is None or spec.loader is None:
        raise SystemExit("failed to load nightly_build_candidates module")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.compute_corr_map, module.enrich_dashboard_columns


def parse_date(value: str | None) -> dt.date:
    if not value:
        raise ValueError("end_date is required (YYYYMMDD or YYYY-MM-DD)")
    for fmt in ("%Y%m%d", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"invalid date: {value}")


def run_compute_coeffs(
    repo_root: Path,
    candidates: Path,
    output_path: Path,
    history_days: int,
    end_date: dt.date,
) -> None:
    cmd = [
        sys.executable,
        "tools/compute_dashboard_coeffs.py",
        "--codes-file",
        str(candidates),
        "--history-days",
        str(history_days),
        "--end-date",
        end_date.strftime("%Y%m%d"),
        "--output",
        str(output_path),
    ]
    subprocess.run(cmd, check=True, cwd=repo_root)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--candidates", required=True, help="Input candidates_nextday CSV")
    ap.add_argument("--output", default="", help="Output CSV (default: in-place)")
    ap.add_argument("--end-date", required=True, help="Anchor date YYYYMMDD or YYYY-MM-DD")
    ap.add_argument("--coeff-history-days", type=int, default=60)
    ap.add_argument("--corr-history-days", type=int, default=180)
    ap.add_argument("--coeff-file", default="", help="Optional precomputed coeffs CSV")
    args = ap.parse_args()

    repo_root = Path(".").resolve()
    compute_corr_map, enrich_dashboard_columns = load_nightly_helpers(repo_root)
    input_path = Path(args.candidates)
    if not input_path.exists():
        raise SystemExit(f"missing candidates file: {input_path}")
    output_path = Path(args.output) if args.output else input_path

    # Ensure output exists before enrichment
    if output_path != input_path:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(input_path, output_path)

    end_date = parse_date(args.end_date)

    coeff_path = Path(args.coeff_file) if args.coeff_file else (
        repo_root
        / "output"
        / "excel"
        / f"dashboard_coeffs_latest_{end_date:%Y%m%d}.csv"
    )
    coeff_path.parent.mkdir(parents=True, exist_ok=True)
    if not coeff_path.exists():
        run_compute_coeffs(
            repo_root,
            output_path,
            coeff_path,
            args.coeff_history_days,
            end_date,
        )

    df = pd.read_csv(output_path)
    tickers = df["Ticker"].dropna().astype(str).tolist() if "Ticker" in df.columns else []
    corr_map = compute_corr_map(
        tickers,
        end_date=end_date,
        lookback_days=int(args.corr_history_days),
    ) if tickers else {}

    enrich_dashboard_columns(output_path, coeff_path, corr_map=corr_map)
    print(f"updated={output_path} coeffs={coeff_path} tickers={len(tickers)}")


if __name__ == "__main__":
    main()
