# ASAGAKE Revalidation Report (march20260320_0325)

- candidates: `C:\AI\asagake\output\excel\candidates_for_20260320.csv`
- window: `2026-03-20` to `2026-03-25`

## Scenario Summary

| scenario | trades | pnl_yen_sum | pnl_bp_mean | win_rate | pf |
| --- | --- | --- | --- | --- | --- |
| actual_like_all | 17 | -847331.4529498604 | -49.06001550757438 | 0.11764705882352941 | 0.02330075092919919 |
| actual_like_no_live_base | 10 | -546937.3471566569 | -53.36261578355609 | 0.1 | 0.001899434455208882 |
| actual_like_live_strong | 4 | -373499.024318502 | -46.68737803981276 | 0.0 | 0.0 |
| actual_like_live_base_only | 7 | -300394.10579320346 | -42.91344368474335 | 0.14285714285714285 | 0.05999862039123666 |
| baseline_current_all | 17 | 95730.3610979607 | 3.862084898277896 | 0.6470588235294118 | 1.3438884867748777 |
| baseline_current_live_strong | 6 | 126098.34441394472 | 10.508195367828725 | 0.6666666666666666 | 2.4763533196545313 |

## Readout

- baseline_current_all pnl: `95,730 yen`
- actual_like_all pnl: `-847,331 yen`
- delta(actual_like - baseline): `-943,062 yen`
- actual_like_live_strong pnl: `-373,499 yen`
- actual_like_live_base_only pnl: `-300,394 yen`

## Daily Summary

| date | trades | pnl_yen | pnl_bp_mean | scenario | diag_missing_intraday | cooldown_minutes | max_trades_per_ticker | stop_after_loss | diag_skip_gapban | diag_skip_trend_mismatch | diag_no_signal | LIVE_STRONG_trades | LIVE_STRONG_pnl_yen | LIVE_STRONG_pnl_bp_mean | LIVE_BASE_trades | LIVE_BASE_pnl_yen | LIVE_BASE_pnl_bp_mean | DEMO_ONLY_trades | DEMO_ONLY_pnl_yen | DEMO_ONLY_pnl_bp_mean |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 2026-03-20 | 0 | 0.0 | 0.0 | actual_like_all | 27.0 | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan |
| 2026-03-23 | 5 | -323584.310692125 | -69.20273592866054 | actual_like_all | nan | nan | nan | nan | 14.0 | 4.0 | 4.0 | nan | nan | nan | nan | nan | nan | nan | nan | nan |
| 2026-03-24 | 4 | -196899.75761771837 | -55.12857373933265 | actual_like_all | nan | nan | nan | nan | 7.0 | 13.0 | 3.0 | nan | nan | nan | nan | nan | nan | nan | nan | nan |
| 2026-03-25 | 8 | -326847.38464001694 | -33.43653612851637 | actual_like_all | nan | nan | nan | nan | 8.0 | 9.0 | 2.0 | nan | nan | nan | nan | nan | nan | nan | nan | nan |
| 2026-03-20 | 0 | 0.0 | 0.0 | actual_like_live_base_only | 12.0 | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan |
| 2026-03-23 | 1 | -51973.81401685882 | -51.973814016858825 | actual_like_live_base_only | nan | nan | nan | nan | 8.0 | 1.0 | 2.0 | nan | nan | nan | nan | nan | nan | nan | nan | nan |
| 2026-03-24 | 2 | -130501.40196152279 | -65.2507009807614 | actual_like_live_base_only | nan | nan | nan | nan | 4.0 | 4.0 | 2.0 | nan | nan | nan | nan | nan | nan | nan | nan | nan |
| 2026-03-25 | 4 | -117918.8898148218 | -29.47972245370545 | actual_like_live_base_only | nan | nan | nan | nan | 3.0 | 3.0 | 2.0 | nan | nan | nan | nan | nan | nan | nan | nan | nan |
| 2026-03-20 | 0 | 0.0 | 0.0 | actual_like_live_strong | 6.0 | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan |
| 2026-03-23 | 1 | -166120.75181605903 | -83.06037590802951 | actual_like_live_strong | nan | nan | nan | nan | 3.0 | 2.0 | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan |
| 2026-03-24 | 1 | -28522.5455443889 | -14.261272772194452 | actual_like_live_strong | nan | nan | nan | nan | 1.0 | 4.0 | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan |
| 2026-03-25 | 2 | -178855.7269580541 | -44.71393173951353 | actual_like_live_strong | nan | nan | nan | nan | 2.0 | 2.0 | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan |
| 2026-03-20 | 0 | 0.0 | 0.0 | actual_like_no_live_base | 15.0 | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan |
| 2026-03-23 | 4 | -271610.4966752662 | -73.50996640661097 | actual_like_no_live_base | nan | nan | nan | nan | 6.0 | 3.0 | 2.0 | nan | nan | nan | nan | nan | nan | nan | nan | nan |
| 2026-03-24 | 2 | -66398.35565619558 | -45.00644649790391 | actual_like_no_live_base | nan | nan | nan | nan | 3.0 | 9.0 | 1.0 | nan | nan | nan | nan | nan | nan | nan | nan | nan |
| 2026-03-25 | 4 | -208928.49482519517 | -37.39334980332729 | actual_like_no_live_base | nan | nan | nan | nan | 5.0 | 6.0 | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan |
| 2026-03-20 | 0 | 0.0 | 0.0 | baseline_current_all | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan |
| 2026-03-23 | 4 | 110610.93272563201 | 12.745469583326672 | baseline_current_all | nan | 5.0 | 2.0 | False | 12.0 | 5.0 | 4.0 | 2.0 | 119258.10878465063 | 29.814527196162654 | 2.0 | -8647.176059018617 | -4.32358802950931 | 0.0 | 0.0 | 0.0 |
| 2026-03-24 | 6 | 75590.20425973613 | 13.11572548205072 | baseline_current_all | nan | 5.0 | 2.0 | False | 7.0 | 12.0 | 2.0 | 2.0 | 37597.943947798165 | 9.39948598694954 | 3.0 | 16089.139705470672 | 5.363046568490225 | 1.0 | 21903.12060646728 | 43.80624121293456 |
| 2026-03-25 | 7 | -90470.77588740742 | -9.145826850698109 | baseline_current_all | nan | 5.0 | 2.0 | False | 6.0 | 9.0 | 3.0 | 2.0 | -30757.708318504076 | -7.68942707962602 | 4.0 | -70784.20134217196 | -17.696050335542992 | 1.0 | 11071.133773268622 | 22.142267546537244 |
| 2026-03-20 | 0 | 0.0 | 0.0 | baseline_current_live_strong | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan | nan |
| 2026-03-23 | 2 | 119258.10878465063 | 29.814527196162654 | baseline_current_live_strong | nan | 5.0 | 2.0 | False | 3.0 | 1.0 | nan | 2.0 | 119258.10878465063 | 29.814527196162654 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 |
| 2026-03-24 | 2 | 37597.943947798165 | 9.39948598694954 | baseline_current_live_strong | nan | 5.0 | 2.0 | False | 1.0 | 3.0 | nan | 2.0 | 37597.943947798165 | 9.39948598694954 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 |
| 2026-03-25 | 2 | -30757.708318504076 | -7.68942707962602 | baseline_current_live_strong | nan | 5.0 | 2.0 | False | 2.0 | 2.0 | nan | 2.0 | -30757.708318504076 | -7.68942707962602 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 | 0.0 |

## Class Breakdown

| scenario | live_demo_class | trades | pnl_yen | pnl_bp_mean |
| --- | --- | --- | --- | --- |
| actual_like_all | LIVE_STRONG | 4 | -373499.024318502 | -46.68737803981276 |
| actual_like_all | LIVE_BASE | 7 | -300394.1057932034 | -42.91344368474334 |
| actual_like_all | DEMO_ONLY | 6 | -173438.32283815494 | -57.812774279384975 |
| actual_like_live_base_only | LIVE_BASE | 7 | -300394.1057932034 | -42.91344368474334 |
| actual_like_live_strong | LIVE_STRONG | 4 | -373499.024318502 | -46.68737803981276 |
| actual_like_no_live_base | LIVE_STRONG | 4 | -373499.024318502 | -46.68737803981276 |
| actual_like_no_live_base | DEMO_ONLY | 6 | -173438.32283815494 | -57.812774279384975 |
| baseline_current_all | LIVE_BASE | 9 | -63342.2376957199 | -7.0380264106355455 |
| baseline_current_all | DEMO_ONLY | 2 | 32974.2543797359 | 32.9742543797359 |
| baseline_current_all | LIVE_STRONG | 6 | 126098.34441394472 | 10.508195367828725 |
| baseline_current_live_strong | LIVE_STRONG | 6 | 126098.34441394472 | 10.508195367828725 |

## Worst Actual-like Trades

| scenario | date | code | session | signal_mode | side | exit_reason | pnl_yen | pnl_bp |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| actual_like_all | 2026-03-23 | 7012.T | AM1030 | j-cross | BUY | SL_SAME_BAR | -166120.75181605903 | -83.06037590802951 |
| actual_like_live_strong | 2026-03-23 | 7012.T | AM1030 | j-cross | BUY | SL_SAME_BAR | -166120.75181605903 | -83.06037590802951 |
| actual_like_no_live_base | 2026-03-23 | 7012.T | AM1030 | j-cross | BUY | SL_SAME_BAR | -166120.75181605903 | -83.06037590802951 |
| actual_like_all | 2026-03-25 | 7012.T | AM1030 | j-cross | BUY | SL_SAME_BAR | -124454.46613195812 | -62.227233065979064 |
| actual_like_live_strong | 2026-03-25 | 7012.T | AM1030 | j-cross | BUY | SL_SAME_BAR | -124454.46613195812 | -62.227233065979064 |
| actual_like_no_live_base | 2026-03-25 | 7012.T | AM1030 | j-cross | BUY | SL_SAME_BAR | -124454.46613195812 | -62.227233065979064 |
| actual_like_all | 2026-03-24 | 1605.T | AM0930 | j-only | SELL | SL_SAME_BAR | -97821.57809006788 | -97.82157809006787 |
| actual_like_live_base_only | 2026-03-24 | 1605.T | AM0930 | j-only | SELL | SL_SAME_BAR | -97821.57809006788 | -97.82157809006787 |
| actual_like_live_base_only | 2026-03-25 | 6981.T | AM0930 | j-only | BUY | SL_SAME_BAR | -56730.11506147715 | -56.73011506147715 |
| actual_like_all | 2026-03-25 | 6981.T | AM0930 | j-only | BUY | SL_SAME_BAR | -56730.11506147715 | -56.73011506147715 |
| actual_like_live_strong | 2026-03-25 | 7267.T | AM1015 | j-cross | BUY | SL_SAME_BAR | -54401.26082609596 | -27.200630413047982 |
| actual_like_all | 2026-03-25 | 7267.T | AM1015 | j-cross | BUY | SL_SAME_BAR | -54401.26082609596 | -27.200630413047982 |
| actual_like_no_live_base | 2026-03-25 | 7267.T | AM1015 | j-cross | BUY | SL_SAME_BAR | -54401.26082609596 | -27.200630413047982 |
| actual_like_all | 2026-03-25 | 1605.T | AM0930 | j-only | BUY | SL_SAME_BAR | -52747.922516085055 | -52.747922516085055 |
| actual_like_live_base_only | 2026-03-25 | 1605.T | AM0930 | j-only | BUY | SL_SAME_BAR | -52747.922516085055 | -52.747922516085055 |

## Best Actual-like Trades

| scenario | date | code | session | signal_mode | side | exit_reason | pnl_yen | pnl_bp |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| actual_like_all | 2026-03-25 | 6723.T | AM0930 | j-only | BUY | TP | 19173.62283952479 | 19.17362283952479 |
| actual_like_live_base_only | 2026-03-25 | 6723.T | AM0930 | j-only | BUY | TP | 19173.62283952479 | 19.17362283952479 |
| actual_like_no_live_base | 2026-03-25 | 4755.T | PM1230 | j-cross | BUY | TIMEOUT | 1040.8486658484667 | 2.081697331696933 |
| actual_like_all | 2026-03-25 | 4755.T | PM1230 | j-cross | BUY | TIMEOUT | 1040.8486658484667 | 2.081697331696933 |
| actual_like_all | 2026-03-23 | 6525.T | MID1030 | j-cross | SELL | SL_SAME_BAR | -27175.225566542264 | -54.35045113308452 |
| actual_like_no_live_base | 2026-03-23 | 6525.T | MID1030 | j-cross | SELL | SL_SAME_BAR | -27175.225566542264 | -54.35045113308452 |
| actual_like_live_base_only | 2026-03-25 | 6501.T | MID1030 | j-cross | BUY | SL | -27614.47507678438 | -27.61447507678438 |
| actual_like_all | 2026-03-25 | 6501.T | MID1030 | j-cross | BUY | SL | -27614.47507678438 | -27.61447507678438 |
| actual_like_no_live_base | 2026-03-24 | 3350.T | AM1030 | j-only | BUY | SL_SAME_BAR | -28522.5455443889 | -14.261272772194452 |
| actual_like_all | 2026-03-24 | 3350.T | AM1030 | j-only | BUY | SL_SAME_BAR | -28522.5455443889 | -14.261272772194452 |
| actual_like_live_strong | 2026-03-24 | 3350.T | AM1030 | j-only | BUY | SL_SAME_BAR | -28522.5455443889 | -14.261272772194452 |
| actual_like_all | 2026-03-25 | 7012.T | AM0945 | j-cross | BUY | SL_SAME_BAR | -31113.61653298953 | -62.227233065979064 |
| actual_like_no_live_base | 2026-03-25 | 7012.T | AM0945 | j-cross | BUY | SL_SAME_BAR | -31113.61653298953 | -62.227233065979064 |
| actual_like_all | 2026-03-24 | 6501.T | MID1030 | j-cross | BUY | SL | -32679.82387145491 | -32.679823871454914 |
| actual_like_live_base_only | 2026-03-24 | 6501.T | MID1030 | j-cross | BUY | SL | -32679.82387145491 | -32.679823871454914 |
