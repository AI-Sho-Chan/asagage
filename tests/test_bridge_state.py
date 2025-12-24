from __future__ import annotations

from pathlib import Path

import pandas as pd

from asagake_io.bridge_state import summarize_execution_events


def test_summarize_execution_events(tmp_path: Path) -> None:
    p = tmp_path / "execution_events_20251224.csv"
    pd.DataFrame(
        [
            {
                "schema_version": "EE.v1",
                "run_id": "R1",
                "event_ts": "2025-12-24T09:01:00+09:00",
                "event_seq": "1",
                "cmd_seq": "2",
                "client_order_id": "O2",
                "broker_order_id": "",
                "exec_event": "ACK",
            },
            {
                "schema_version": "EE.v1",
                "run_id": "R1",
                "event_ts": "2025-12-24T09:01:01+09:00",
                "event_seq": "2",
                "cmd_seq": "2",
                "client_order_id": "O2",
                "broker_order_id": "",
                "exec_event": "ACK",
            },
            {
                "schema_version": "EE.v1",
                "run_id": "R1",
                "event_ts": "2025-12-24T09:01:02+09:00",
                "event_seq": "5",
                "cmd_seq": "7",
                "client_order_id": "O7",
                "broker_order_id": "",
                "exec_event": "REJECT",
            },
        ]
    ).to_csv(p, index=False, encoding="utf-8-sig")

    s = summarize_execution_events(p)
    assert s.run_id == "R1"
    assert s.max_cmd_seq == 7
    assert s.max_event_seq == 5
    assert s.ack_like_duplicate_cmd_seq == 1

