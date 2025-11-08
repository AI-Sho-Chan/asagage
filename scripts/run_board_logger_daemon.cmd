@echo off
cd /d C:\AI\asagake
"C:\Python313\python.exe" scripts\board_logger_daemon.py --board excel\BoardLogger.xlsx --outdir output\board_logs --dashboard C:\AI\asagake\ASAGAKE.xlsm --dash-outdir output\j_logs --interval 5 --retain-days 30
