@echo off
cd /d C:\AI\asagake
C:\Python313\python.exe scripts\nightly_build_candidates.py --excel SHINSOKU.xlsm --excel-summary --universe-mode yahoo-top --universe-size 300 --universe-metric vol

