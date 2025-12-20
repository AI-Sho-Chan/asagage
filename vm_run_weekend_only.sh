#!/bin/bash
set -euo pipefail

LOCK_FILE="/tmp/asagake_weekend_batch.lock"
exec 9>"$LOCK_FILE"
if command -v flock >/dev/null 2>&1; then
  if ! flock -n 9; then
    echo "[vm_run_weekend_only] another weekend batch is already running; exiting." >&2
    exit 0
  fi
fi

LOG_DIR="$HOME/cloud_logs"
mkdir -p "$LOG_DIR"
# Force JST for file naming and date tags, even if cron's environment is UTC.
export TZ="${TZ:-Asia/Tokyo}"
STAMP=$(date +%Y%m%d_%H%M%S)
LOG_FILE="$LOG_DIR/weekend_${STAMP}.log"

# Auto-stop the VM after the batch finishes (success or failure) to minimize cost.
# Set AUTO_SHUTDOWN=0 to keep the VM running for manual debugging.
AUTO_SHUTDOWN="${AUTO_SHUTDOWN:-1}"
cleanup_and_shutdown() {
  local exit_code=$?
  if [ "$AUTO_SHUTDOWN" = "1" ]; then
    echo "[vm_run_weekend_only] exiting (code=${exit_code}); shutting down VM..." >&2
    sudo -n shutdown -h now || true
  fi
}
trap cleanup_and_shutdown EXIT

{
  cd "$HOME/asagage"
  # Keep VM repo up-to-date (best-effort, do not fail the batch on git issues).
  if command -v git >/dev/null 2>&1 && git rev-parse --is-inside-work-tree >/dev/null 2>&1; then
    git fetch origin || true
    # The batch updates tracked state/log files during runs; auto-stash so git pull works reliably.
    if [ -n "$(git status --porcelain)" ]; then
      git stash push -u -m "autostash before pull (weekend batch)" || true
    fi
    git pull --ff-only || true
    # Best-effort restore local batch state after pulling updates.
    if git stash list | grep -q "autostash before pull (weekend batch)"; then
      git stash pop || true
    fi
  fi
  # Activate Python virtualenv for weekend batch
  source "$HOME/asagake-venv/bin/activate"
  TARGET_DATE="${1:-$(date +%Y%m%d)}"

  # Ensure Parquet support exists so 1-minute bars can be saved locally.
  # Without this, the batch tends to re-download data for each plan (slow).
  if ! python -c "import pyarrow" >/dev/null 2>&1; then
    echo "[vm_run_weekend_only] pyarrow not found; installing (one-time)..." >&2
    python -m pip install -q pyarrow || echo "[vm_run_weekend_only] pyarrow install failed; minute cache may not persist" >&2
  fi

  # Non-essential opt30 cache can grow very large; clear it once per weekend run
  # (remove directory itself to avoid slow per-file deletes).
  if [ -d "cache/opt30" ]; then
    rm -rf cache/opt30
  fi
  mkdir -p cache/opt30

  # Pull the "Top200 regulars" 1m cache from GCS (built on Windows) to reduce network fetches on the VM.
  # Best-effort: the batch still runs even if this sync fails.
  REGULARS_CACHE_BUCKET="${REGULARS_CACHE_BUCKET:-gs://asagage-weekend-output/yahoo_1m_regulars}"
  if command -v gsutil >/dev/null 2>&1; then
    echo "[vm_run_weekend_only] syncing regulars 1m cache from ${REGULARS_CACHE_BUCKET} ..." >&2
    gsutil -m rsync -r "${REGULARS_CACHE_BUCKET}" "$HOME/asagage/data/raw/yahoo_1m" || true
  else
    echo "[vm_run_weekend_only] gsutil not found; skip regulars cache sync" >&2
  fi

  # Pull the latest abnormal ticker list (built by Windows DailyReplay) if available.
  ABNORMAL_CODES_OBJECT="${ABNORMAL_CODES_OBJECT:-gs://asagage-weekend-output/universe/abnormal_codes_latest.csv}"
  if command -v gsutil >/dev/null 2>&1; then
    echo "[vm_run_weekend_only] fetching abnormal list from ${ABNORMAL_CODES_OBJECT} ..." >&2
    gsutil cp "${ABNORMAL_CODES_OBJECT}" "$HOME/asagage/data/universe/abnormal_codes_latest.csv" || true
  fi

  # Build weekly Top300 universe (close * weekly volume)
  python tools/build_master_topvol_universe.py --topn 300 --lookback 5 --tag "$TARGET_DATE"
  UNIVERSE_FILE="$HOME/asagage/data/universe/topvol_${TARGET_DATE}.csv"

  # Update "Top200 regulars (ever)" list locally (used by weekend incremental policy).
  # Best-effort: if topvol history is not sufficient yet, the script still runs.
  python tools/build_top_regulars_universe.py --lookback-files 20 --topn 200 --tag "$TARGET_DATE" --update-ever || true

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

  WEEKEND_INCREMENTAL="${WEEKEND_INCREMENTAL:-1}"
  WEEKEND_FORCE_FULL_RESET="${WEEKEND_FORCE_FULL_RESET:-0}"
  EXTRA_WEEKEND_FLAGS=()
  if [ "$WEEKEND_INCREMENTAL" = "1" ]; then
    EXTRA_WEEKEND_FLAGS+=(--weekend-incremental --weekend-monthly-reset)
  fi
  if [ "$WEEKEND_FORCE_FULL_RESET" = "1" ]; then
    EXTRA_WEEKEND_FLAGS+=(--weekend-force-full-reset)
  fi

  # Weekend batch (coarse + refine, longer window, VWAP filter applied inside script)
  python scripts/nightly_build_candidates.py \
    --universe-mode yahoo-top --universe-size 300 \
    --universe-source "$UNIVERSE_FILE" \
    --lookback 60 --chunk-days 5 --train-days 12 --forward-days 4 \
    --min-train-trades 10 --min-forward-trades 2 --forward-pf-min 1.3 \
    --disable-minute-cache \
    --min-forward-winrate 0.60 --gap-guard-abs-bp 80 --gap-guard-dir-bp 40 \
    --slipbp 4 --feebp 4 --liquidity-quantile 0.3 --jobs 16 \
    --run-type weekend --plan-profile weekend \
    --enable-asha --enable-bayes --bayes-trials 60 --bayes-timeout 600 \
    --mask-ineffective --mask-window 20 --mask-threshold 1.05 \
    --enable-rd-windows \
    --enable-market-features --headless --coeff-history-days 5 \
    --target-date "$TARGET_DATE" \
    "${EXTRA_WEEKEND_FLAGS[@]}"

  # Optional weekday-style sanity run on the same universe (smaller size).
  # Disabled by default on the weekend VM because it can create very large opt30 caches.
  if [ "${RUN_WEEKDAY_SANITY:-0}" = "1" ]; then
    python scripts/nightly_build_candidates.py \
      --jobs 12 --universe-mode yahoo-top --universe-size 150 \
      --universe-source "$UNIVERSE_FILE" \
      --run-type weekday --plan-profile weekday \
      --enable-asha --enable-bayes --bayes-trials 32 --bayes-timeout 300 \
      --mask-ineffective --enable-market-features \
      --disable-minute-cache \
      --min-forward-winrate 0.60 --headless --coeff-history-days 5 \
      --target-date "$TARGET_DATE"
  fi

  python scripts/run_trade_analysis.py || true

  # Weekly WF summary report (candidates x trades x expected PnL) and email
  python analysis/build_weekly_wf_report.py \
    --week-ending "$TARGET_DATE" \
    --email \
    --recipient "shouichi.ikeda@gmail.com" || true

  gsutil -m rsync -r "$HOME/asagage/output" "gs://asagage-weekend-output/output"
} >> "$LOG_FILE" 2>&1
