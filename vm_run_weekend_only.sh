#!/bin/bash
set -euo pipefail

LOG_DIR="$HOME/cloud_logs"
mkdir -p "$LOG_DIR"
STAMP=$(date +%Y%m%d_%H%M%S)
LOG_FILE="$LOG_DIR/weekend_${STAMP}.log"

{
  cd "$HOME/asagage"
  # Activate Python virtualenv for weekend batch
  source "$HOME/asagake-venv/bin/activate"
  TARGET_DATE=$(date +%Y%m%d)

  # Build weekly Top300 universe (close * weekly volume)
  python tools/build_master_topvol_universe.py --topn 300 --lookback 5 --tag "$TARGET_DATE"
  UNIVERSE_FILE="$HOME/asagage/data/universe/topvol_${TARGET_DATE}.csv"

  # Extended minute-cache for Top300 (60 history days, 120 backfill, best-effort)
  if [ -f "$UNIVERSE_FILE" ]; then
    python tools/update_minute_cache.py \
      --codes-file "$UNIVERSE_FILE" \
      --universe-glob "data/universe/topvol_*.csv" \
      --universe-limit 800 \
      --history-days 60 \
      --backfill-days 120 \
      --batch-size 120 \
      --pause 0.2 || true
  fi

  # Weekend batch (coarse + refine, longer window, VWAP filter applied inside script)
  python scripts/nightly_build_candidates.py \
    --universe-mode yahoo-top --universe-size 300 \
    --universe-source "$UNIVERSE_FILE" \
    --lookback 60 --chunk-days 5 --train-days 12 --forward-days 4 \
    --min-train-trades 10 --min-forward-trades 2 --forward-pf-min 1.3 \
    --min-forward-winrate 0.60 --gap-guard-abs-bp 80 --gap-guard-dir-bp 40 \
    --slipbp 4 --feebp 4 --liquidity-quantile 0.3 --jobs 16 \
    --run-type weekend --plan-profile weekend \
    --enable-asha --enable-bayes --bayes-trials 60 --bayes-timeout 600 \
    --mask-ineffective --mask-window 20 --mask-threshold 1.05 \
    --cache-refresh-weekend --enable-rd-windows \
    --enable-market-features --headless --coeff-history-days 5 \
    --target-date "$TARGET_DATE"

  # Optional weekday-style sanity run on the same universe (smaller size)
  python scripts/nightly_build_candidates.py \
    --jobs 12 --universe-mode yahoo-top --universe-size 150 \
    --universe-source "$UNIVERSE_FILE" \
    --run-type weekday --plan-profile weekday \
    --enable-asha --enable-bayes --bayes-trials 32 --bayes-timeout 300 \
    --mask-ineffective --enable-market-features \
    --min-forward-winrate 0.60 --headless --coeff-history-days 5 \
    --target-date "$TARGET_DATE"

  python scripts/run_trade_analysis.py || true

  # Weekly WF summary report (candidates × trades × expected PnL) and email
  python analysis/build_weekly_wf_report.py \
    --week-ending "$TARGET_DATE" \
    --email \
    --recipient "shouichi.ikeda@gmail.com" || true

  gsutil -m rsync -r "$HOME/asagage/output" "gs://asagage-weekend-output/output"
} >> "$LOG_FILE" 2>&1


