#!/usr/bin/env python3
from __future__ import annotations

import argparse
import subprocess
from pathlib import Path


def run_local_weekend(universe_size: int, jobs: int) -> int:
    args = [
        "python", "scripts/nightly_build_candidates.py",
        "--excel", "ASAGAKE.xlsm",
        "--base-out", "output/bt30/WEEKLY_LOCAL",
        "--run-type", "weekend", "--plan-profile", "weekend",
        "--universe-mode", "yahoo-top", "--universe-size", str(universe_size),
        "--lookback", "60", "--chunk-days", "5", "--train-days", "12", "--forward-days", "4",
        "--min-train-trades", "12", "--min-forward-trades", "5", "--forward-pf-min", "1.3", "--min-forward-ci", "0.65",
        "--gap-guard-abs-bp", "80", "--gap-guard-dir-bp", "40",
        "--liquidity-quantile", "0.3", "--jobs", str(jobs),
        "--enable-asha", "--enable-bayes", "--bayes-trials", "20", "--bayes-timeout", "600",
        "--mask-ineffective", "--mask-window", "20", "--mask-threshold", "1.05",
        "--enable-market-features", "--excel-summary", "--analysis-ledger",
    ]
    return subprocess.call(args)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--backend", choices=["local", "github"], default="local")
    ap.add_argument("--universe-size", type=int, default=150)
    ap.add_argument("--jobs", type=int, default=8)
    args = ap.parse_args()

    if args.backend == "local":
        code = run_local_weekend(args.universe_size, args.jobs)
        raise SystemExit(code)
    else:
        wf = Path(".github/workflows/weekend.yml")
        if not wf.exists():
            print("GitHub Actions workflow missing: .github/workflows/weekend.yml")
            raise SystemExit(1)
        print("Push to GitHub and dispatch the 'Weekend Screening (Cloud)' workflow.")
        print("Repository: actions -> Weekend Screening (Cloud) -> Run workflow で実行してください。")


if __name__ == "__main__":
    main()

