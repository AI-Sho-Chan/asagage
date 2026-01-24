#!/bin/bash
set -euo pipefail

# Weekend batch safety net:
# - Triggered by cron "@reboot"
# - If the VM starts late and misses the fixed 16:30 JST cron,
#   this script starts the weekend batch once (best-effort).
#
# Notes:
# - Cron uses VM local timezone (Asia/Tokyo).
# - The main batch script has its own flock lock + auto-shutdown, so double-start is safe.

export TZ="${TZ:-Asia/Tokyo}"

NOW_DOW="$(date +%u)" # 1=Mon ... 5=Fri
NOW_HHMM="$(date +%H%M)"
TARGET_DATE="$(date +%Y%m%d)"

# Only do anything on Fridays.
if [ "${NOW_DOW}" != "5" ]; then
  exit 0
fi

# Only trigger after the normal 16:30 JST cron window.
# If the VM starts too early, let the normal 16:30 cron handle it.
#
# We allow the late-start window to extend through the end of Friday.
# This covers cases where the zone is temporarily resource-exhausted and
# the VM can only start later in the evening.
if [ "${NOW_HHMM}" -lt "1630" ] || [ "${NOW_HHMM}" -gt "2359" ]; then
  exit 0
fi

LOG_DIR="${HOME}/cloud_logs"
mkdir -p "${LOG_DIR}"

# If a weekend log for this date already exists, assume the batch already started (or completed).
if ls "${LOG_DIR}/weekend_${TARGET_DATE}_"*.log >/dev/null 2>&1; then
  echo "[weekend_bootstrap] weekend log already exists for ${TARGET_DATE}; skip." >> "${LOG_DIR}/cron_stdout.log"
  exit 0
fi

echo "[weekend_bootstrap] VM boot detected after 16:30 JST; starting weekend batch for ${TARGET_DATE}..." >> "${LOG_DIR}/cron_stdout.log"

# Run in background so @reboot does not block.
nohup bash "${HOME}/asagage/vm_run_weekend_only.sh" "${TARGET_DATE}" >> "${LOG_DIR}/cron_stdout.log" 2>&1 &
