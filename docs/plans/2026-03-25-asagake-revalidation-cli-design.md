# ASAGAKE Revalidation CLI Design

Date: 2026-03-25
Branch: codex/asagake-revalidation-cli

## Goal

Build a small Python CLI that re-tests the VWAP reversion hypothesis with rules closer to the Excel DEMO lifecycle than the current DailyReplay.

## Non-goals

- No Excel UI rebuild
- No Gmail integration
- No production/live execution
- No attempt to perfectly reconstruct RSS bid/ask tick data

## Why this exists

The current evidence mixes four things:

1. hypothesis quality
2. parameter/candidate quality
3. Excel DEMO implementation bugs
4. replay/actual mismatch

This tool isolates (1) and (2) better by replacing the current replay with an "actual-like" replay.

## Inputs

- Candidates CSV:
  - default: `C:\AI\asagake\output\excel\candidates_nextday.csv`
  - optional explicit file path
- Minute data root:
  - default: `C:\AI\asagake\data\raw\yahoo_1m`
- Date range:
  - explicit start/end dates

## Core simulation model

For each candidate row:

1. Load 1-minute OHLCV
2. Compute intraday VWAP and ATR proxy
3. Apply dynamic J threshold:
   - `adjJth = J_th + BiasSlope_row*(Bias_bp/100) + GapSlope_row*|Gap_bp|/100 + CorrSlope_row*Corr(driver)*(Bias_bp/100)`
4. Find first signal in the session:
   - `j-only`: `abs(J) >= adjJth`
   - `j-cross`: first crossing into `abs(J) >= adjJth`
5. At the signal bar, compute actual-like passive limit:
   - base price = VWAP, else PrevClose, else current close
   - BUY limit = `basePrice - 0.001 * abs(adjJth) * basePrice`
   - SELL limit = `basePrice + 0.001 * abs(adjJth) * basePrice`
6. Approximate fill from minute bars:
   - BUY fills if a later bar low touches the limit
   - SELL fills if a later bar high touches the limit
7. After fill, manage:
   - TP/SL from candidate row coefficients
   - 30-minute timeout
   - end-of-day fallback
   - if TP and SL touch in the same bar, choose conservative SL-first

## Portfolio / scheduling rules

- Global daily entry cap: 20
- One open position per ticker at a time
- Candidate class analysis:
  - ALL
  - LIVE_STRONG only
  - NO_LIVE_BASE
- Use the same candidate snapshot across the inspected date window unless an explicit candidates file is passed

## Outputs

- Trade detail CSV
- Daily summary CSV
- Class summary CSV
- Markdown report

## Validation questions

1. Does actual-like replay remain negative on the recent window?
2. Does LIVE_STRONG survive while LIVE_BASE collapses?
3. Does moving from close-entry replay to actual-like passive-entry replay materially worsen results?
4. Is the hypothesis still alive in a narrow slice, or is it weak even after cleanup?

## Success / failure interpretation

- If actual-like replay is negative even for LIVE_STRONG:
  - hypothesis is probably too weak to continue
- If current replay is weak but actual-like replay is much worse:
  - implementation/execution translation is a major failure point
- If LIVE_STRONG remains positive while LIVE_BASE is negative:
  - the idea may only survive as a much narrower strategy

