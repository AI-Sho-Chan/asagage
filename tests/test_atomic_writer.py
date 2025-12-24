from __future__ import annotations

import csv
from pathlib import Path

from asagake_io.atomic_writer import atomic_write_csv
from asagake_io.csv_schemas import OC_V1, schema_columns


def test_atomic_write_csv_repeated_updates_leave_no_tmp(tmp_path: Path) -> None:
    out = tmp_path / "orders_cmd_20251224.csv"
    cols = schema_columns(OC_V1)

    # Seed with one file so "reader during update" always sees a valid CSV.
    atomic_write_csv(
        out,
        columns=cols,
        rows=[
            {
                "schema_version": OC_V1,
                "run_id": "RUN_TEST",
                "cmd_ts": "2025-12-24T09:00:00+09:00",
                "cmd_seq": 1,
                "action": "PLACE",
                "ticker": "7203",
                "side": "BUY",
                "qty": 100,
                "order_type": "LIMIT",
                "limit_price": 2800.0,
                "time_in_force": "DAY",
                "candidate_id": "C_7203_AM0930_abcdef",
                "decision_id": "D_20251224_090000_0001",
                "client_order_id": "O_20251224_090000_0001",
            }
        ],
    )

    for i in range(2, 102):
        atomic_write_csv(
            out,
            columns=cols,
            rows=[
                {
                    "schema_version": OC_V1,
                    "run_id": "RUN_TEST",
                    "cmd_ts": f"2025-12-24T09:00:{i:02d}+09:00",
                    "cmd_seq": i,
                    "action": "PLACE",
                    "ticker": "7203",
                    "side": "BUY",
                    "qty": 100,
                    "order_type": "LIMIT",
                    "limit_price": 2800.0 + i / 10.0,
                    "time_in_force": "DAY",
                    "candidate_id": "C_7203_AM0930_abcdef",
                    "decision_id": "D_20251224_090000_0001",
                    "client_order_id": "O_20251224_090000_0001",
                }
            ],
        )

        # Must always be parseable as CSV (simulates "reader during update").
        with open(out, "r", encoding="utf-8-sig", newline="") as f:
            r = csv.reader(f)
            header = next(r)
            assert header == cols
            row = next(r)
            assert len(row) == len(header)

        # No temp files should remain.
        assert not list(tmp_path.glob("orders_cmd_20251224.csv.*.tmp"))

