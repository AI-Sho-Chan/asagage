from __future__ import annotations

from pathlib import Path

import pandas as pd

from asagake_io.csv_schemas import DT_V1
from asagake_io.validator import validate_csv
from tools.bridge_shadow_engine import ShadowConfig, run_shadow_engine


def test_shadow_engine_writes_dt_once(tmp_path: Path) -> None:
    base = tmp_path
    date_tag = "20251224"

    (base / "output" / "excel" / "outbox").mkdir(parents=True, exist_ok=True)
    (base / "output" / "excel" / "inbox").mkdir(parents=True, exist_ok=True)
    (base / "output" / "excel").mkdir(parents=True, exist_ok=True)
    (base / "analysis").mkdir(parents=True, exist_ok=True)

    candidates_path = base / "output" / "excel" / "candidates_nextday.csv"
    snapshots_path = base / "output" / "excel" / "outbox" / f"market_snapshots_{date_tag}.csv"
    dt_path = base / "analysis" / f"decision_trace_{date_tag}.csv"
    oc_path = base / "output" / "excel" / "inbox" / f"orders_cmd_{date_tag}.csv"

    pd.DataFrame(
        [
            {"Ticker": "7203.T", "session": "AM0930", "J_th": "1.2", "NoTradeMin": "5", "GapBanPct": "1.0"},
        ]
    ).to_csv(candidates_path, index=False, encoding="utf-8-sig")

    pd.DataFrame(
        [
            {
                "schema_version": "MS.v1",
                "run_id": "R1",
                "snap_ts": "2025-12-24T09:10:00+09:00",
                "ticker": "7203.T",
                "last": "100",
                "bid": "99",
                "ask": "101",
                "vwap": "100",
                "prev_close": "98",
                "data_quality": "OK",
            }
        ]
    ).to_csv(snapshots_path, index=False, encoding="utf-8-sig")

    cfg = ShadowConfig(
        date_tag=date_tag,
        base_dir=base,
        run_id="20251224_DEMO_PC01_001",
        engine_version="test",
        emit_orders=False,
        once=True,
        follow_seconds=0,
        poll_interval_sec=0.01,
        max_orders_per_tick=1,
        default_qty=100,
    )

    run_shadow_engine(
        config=cfg,
        candidates_path=candidates_path,
        snapshots_path=snapshots_path,
        decision_trace_path=dt_path,
        orders_cmd_path=oc_path,
    )

    assert dt_path.exists()
    errs = validate_csv(dt_path, schema_version=DT_V1)
    assert not errs

