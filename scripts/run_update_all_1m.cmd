@echo off
cd /d C:\AI\asagake
"C:\Python313\python.exe" tools\update_minute_cache.py ^
  --universe-glob data/universe/topvol_*.csv ^
  --universe-limit 1000 ^
  --codes-file output/excel/candidates_nextday.csv ^
  --history-days 1 ^
  --backfill-days 0
