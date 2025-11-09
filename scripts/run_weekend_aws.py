#!/usr/bin/env python3
from __future__ import annotations

import argparse
import subprocess
from pathlib import Path


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--cluster", default="infra/aws/ray-cluster.yaml")
    ap.add_argument("--universe-size", type=int, default=150)
    ap.add_argument("--jobs", type=int, default=48)
    args = ap.parse_args()

    cluster = Path(args.cluster)
    if not cluster.exists():
        print(f"cluster file not found: {cluster}")
        raise SystemExit(1)

    # 1) Up cluster (idempotent)
    subprocess.check_call(["ray", "up", str(cluster)])

    # 2) Sync repo
    subprocess.check_call(["ray", "rsync_up", str(cluster), ".", "/home/ec2-user/asagake"])

    # 3) Run weekend on head
    cmd = (
        "cd asagake && python scripts/nightly_build_candidates.py "
        f"--excel ASAGAKE.xlsm --base-out output/bt30/WEEKLY_AWS "
        f"--run-type weekend --plan-profile weekend --universe-mode yahoo-top --universe-size {args.universe_size} "
        "--lookback 60 --chunk-days 5 --train-days 12 --forward-days 4 "
        "--min-train-trades 12 --min-forward-trades 5 --forward-pf-min 1.3 --min-forward-winrate 0.60 "
        "--gap-guard-abs-bp 80 --gap-guard-dir-bp 40 --liquidity-quantile 0.3 "
        f"--jobs {args.jobs} --enable-asha --enable-bayes --bayes-trials 20 --bayes-timeout 600 "
        "--mask-ineffective --mask-window 20 --mask-threshold 1.05 --enable-market-features --excel-summary --analysis-ledger"
    )
    subprocess.check_call(["ray", "exec", str(cluster), cmd])

    # 4) Pull output back
    subprocess.check_call(["ray", "rsync_down", str(cluster), "/home/ec2-user/asagake/output", "./output"])

    print("weekend run on AWS completed. Review ./output/bt30/WEEKLY_AWS ...")


if __name__ == "__main__":
    main()

