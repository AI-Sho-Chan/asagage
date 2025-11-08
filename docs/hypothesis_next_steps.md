# Hypothesis Follow-up Plan (2025-10-30)

## Hypothesis 1 / 2 (市場トレンド・ギャップ連動)
- Prepare two comparison runs:
  1. Baseline (current settings) for reference.
  2. Market-adjusted: --market-threshold-up-bp 20 / --market-threshold-down-bp 20, plus provisional gap rules (e.g. +0.2 to J_th when gap >=120bp).
- Capture outputs in output/bt30_test/HYPOTHESIS_20251030_AM09_50_marketX. Compare market_bias buckets and winrate deltas via eports/param_stats.

## Hypothesis 3 (EntryPrice フォールバック)
- Verify RSS data availability per ticker; if missing, retain formula but log to Trace.
- Draft fallback routine: when RSS returns空欄, use PreOpenMid as substitute so EntryBuy/SellPx remain populated.
- Test via Queue Now dry-run to ensure Orders sheet captures demo entries.

## H1/H2/H3 (Volume Spike / Repeat / Sector Density)
- From hypothesis_trade_ledger.csv, compute quantiles for ol_spike, epeat_index, sector_density.
- Define candidate thresholds (e.g. vol_spike >=3.0, repeat >=3, density >=3) for masking.
- Update 	ools/update_ineffective_bands.py and state/optuna_priors.json workflow to accept these filters once validated.
