# Bridge v1 integration test (Excel DEMO)

This document fixes the manual “smoke test” steps to validate Bridge v1 end-to-end without breaking the existing workflow
(`output/excel/candidates_nextday.csv` + Orders sheet).

## Preconditions
- Windows machine with `C:\AI\asagake` checked out
- `ASAGAKE.xlsm` has the latest VBA module applied via `scripts/update_asagake_vba.ps1`
- Excel is allowed to read/write under `C:\AI\asagake\output\excel\`

## 1) Generate one OC.v1 command (Python → Excel inbox)
From PowerShell:

`cd C:\AI\asagake`

Pick a date tag (JST): `YYYYMMDD` and a run id (string).

Example:

`python tools/bridge_smoketest_orders.py --date 20251224 --run-id 20251224_DEMO_PC01_001 --ticker 7203.T --side BUY --qty 100 --limit-price 2800`

Expected:
- `output/excel/inbox/orders_cmd_20251224.csv` is created/updated (UTF-8 BOM) and contains a single row with `cmd_seq`.

## 2) Start Excel DEMO and let it consume the inbox command
In Excel:
1. Open `C:\AI\asagake\ASAGAKE.xlsm`
2. Ensure RSS is connected if needed (DEMO is OK without fills)
3. Press `Demo Start` (RunStatus becomes `DEMO_RUNNING`)
4. Wait 5–10 seconds (AutoTickV2 interval is 5s)

Expected:
- The `Orders` sheet gets a new row with:
  - `mode=bridge_cmd`
  - `status=ORDERED`
  - `note` contains `BRIDGE cmd_seq=...`

## 3) Check outbox files (Excel → Python outbox)
Expected files (created on first tick, then append-only):
- `output/excel/outbox/market_snapshots_YYYYMMDD.csv`
- `output/excel/outbox/execution_events_YYYYMMDD.csv`

## 4) Validate CSVs (Python)
From PowerShell:

`python -c "import sys; from pathlib import Path; sys.path.insert(0, str(Path('src').resolve())); from asagake_io.validator import validate_csv; from asagake_io.csv_schemas import MS_V1, EE_V1, OC_V1; print(validate_csv(Path('output/excel/outbox/market_snapshots_20251224.csv'), schema_version=MS_V1)); print(validate_csv(Path('output/excel/outbox/execution_events_20251224.csv'), schema_version=EE_V1)); print(validate_csv(Path('output/excel/inbox/orders_cmd_20251224.csv'), schema_version=OC_V1))"`

Expected:
- Each `validate_csv(...)` prints an empty list `[]`.

## 5) Regression: DailyReplay output must not change
Run the same day twice (with and without decision trace enabled) and compare:
- `analysis/daily_trades_YYYYMMDD.csv` row count
- sum of `pnl_yen` / `pnl_bp` columns (whatever the script reports)

DecisionTrace output must be additive only; it must not change the trading results.

