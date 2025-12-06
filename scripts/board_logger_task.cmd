@echo off
cd /d C:\AI\asagake
C:\Python313\python.exe scripts/board_logger_daemon.py --dashboard "C:\AI\asagake\ASAGAKE.xlsm" --dash-outdir "C:\AI\asagake\output\j_logs" --interval 5 --retain-days 30

