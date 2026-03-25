# ASAGAKE Archive Decision

- Status: `ARCHIVED`
- Decision date: `2026-03-25`
- Scope: stop this project as an active trading-system development effort

## Why It Was Archived

The final revalidation used a small Python CLI that tried to match the actual DEMO behavior more closely than the legacy replay.

- Window: `2026-03-20` to `2026-03-25`
- Baseline replay (`baseline_current_all`): `+95,730 yen`
- Actual-like replay (`actual_like_all`): `-847,331 yen`
- Delta versus baseline: `-943,062 yen`
- Actual-like `LIVE_STRONG`: `-373,499 yen`

This matters because the old replay looked mildly profitable, but the more realistic replay flipped strongly negative. That means the earlier validation path was too optimistic.

## Main Findings

1. The core question was not only the VWAP mean-reversion idea itself. The bigger problem was that the old replay, the Excel execution path, and the DEMO lifecycle were not aligned.
2. After we moved the replay closer to actual execution assumptions, the apparent edge disappeared.
3. The clean recent actual sample and the actual-like replay both showed similar behavior: mostly BUY-side trades that hit SL quickly.
4. Because of that, a large Excel rebuild is not justified.

## Archive Rule

- No further feature work on the current XLSM trading system.
- No live deployment.
- No additional optimization on top of the current rule set.

## If This Is Ever Reopened

Reopen only if all of the following are true.

1. A new, narrower hypothesis is written first.
2. Validation uses one small reproducible engine, not mixed Excel and replay rules.
3. The new validation proves positive across multiple recent weeks before any system rebuild starts.

## Related Evidence

- `reports/revalidation_report_march20260320_0325.md`
- `reports/revalidation_report_feb20260209_0218.md`
