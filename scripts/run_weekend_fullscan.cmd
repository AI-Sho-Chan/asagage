@echo off
set PYTHON=python
set REPO=C:\AI\asagake
cd /d %REPO%
%PYTHON% scripts\nightly_build_candidates.py --universe-mode yahoo-top --universe-size 300 --lookback 60 --chunk-days 5 --train-days 12 --forward-days 4 --min-train-trades 10 --min-forward-trades 2 --forward-pf-min 1.3 --gap-guard-abs-bp 80 --gap-guard-dir-bp 40 --slipbp 4 --feebp 4 --liquidity-quantile 0.5 --jobs 0 --run-type weekend --enable-asha --enable-bayes --bayes-trials 60 --mask-ineffective --mask-window 20 --mask-threshold 1.05 --cache-refresh-weekend --enable-rd-windows
