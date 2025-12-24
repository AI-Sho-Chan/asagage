from __future__ import annotations

from pathlib import Path

import pandas as pd

from tools.bridge_healthcheck import run_healthcheck


def _write_csv(path: Path, header: list[str], rows: list[list[object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(rows, columns=header)
    df.to_csv(path, index=False, encoding="utf-8-sig")


def _write_bytes(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)


def test_healthcheck_ok_with_partial_ack(tmp_path: Path) -> None:
    base = tmp_path
    date_tag = "20251224"

    outbox = base / "output" / "excel" / "outbox"
    inbox = base / "output" / "excel" / "inbox"

    _write_csv(
        outbox / f"market_snapshots_{date_tag}.csv",
        [
            "schema_version",
            "run_id",
            "snap_ts",
            "ticker",
            "last",
            "bid",
            "ask",
            "vwap",
            "cum_volume",
            "prev_close",
            "data_quality",
        ],
        [
            ["MS.v1", "R1", "2025-12-24T09:00:00+09:00", "7203", "100", "99", "101", "100", "", "98", "OK"],
            ["MS.v1", "R1", "2025-12-24T09:00:05+09:00", "7203", "100", "99", "101", "100", "", "98", "OK"],
        ],
    )

    _write_csv(
        inbox / f"orders_cmd_{date_tag}.csv",
        [
            "schema_version",
            "run_id",
            "cmd_ts",
            "cmd_seq",
            "action",
            "ticker",
            "side",
            "qty",
            "order_type",
            "limit_price",
            "time_in_force",
            "candidate_id",
            "decision_id",
            "client_order_id",
        ],
        [
            ["OC.v1", "R1", "2025-12-24T09:01:00+09:00", "1", "PLACE", "7203", "BUY", "100", "LIMIT", "100", "DAY", "C1", "D1", "O1"],
            ["OC.v1", "R1", "2025-12-24T09:01:01+09:00", "2", "PLACE", "7203", "BUY", "100", "LIMIT", "100", "DAY", "C1", "D2", "O2"],
        ],
    )

    _write_csv(
        outbox / f"execution_events_{date_tag}.csv",
        [
            "schema_version",
            "run_id",
            "event_ts",
            "event_seq",
            "cmd_seq",
            "client_order_id",
            "broker_order_id",
            "exec_event",
            "ticker",
            "side",
        ],
        [
            ["EE.v1", "R1", "2025-12-24T09:01:00+09:00", "1", "1", "O1", "", "ACK", "7203", "BUY"],
        ],
    )

    result, _ = run_healthcheck(date_tag=date_tag, base_dir=base)
    assert result.exit_code == 0
    assert result.csv_path.exists()
    assert result.txt_path.exists()


def test_healthcheck_recovers_market_snapshots_variant_prefix(tmp_path: Path) -> None:
    base = tmp_path
    date_tag = "20251224"

    outbox = base / "output" / "excel" / "outbox"
    inbox = base / "output" / "excel" / "inbox"

    ms_text = (
        "schema_version,run_id,snap_ts,ticker,last,bid,ask,vwap,cum_volume,prev_close,data_quality\n"
        "MS.v1,R1,2025-12-24T09:00:00+09:00,7203,100,99,101,100,,98,OK\n"
    )
    variant_prefix = bytes.fromhex("11 20 01 00 DC 00 00 00 00 00 00 00")
    ms_bytes = variant_prefix + b"\xEF\xBB\xBF" + ms_text.encode("utf-8")
    _write_bytes(outbox / f"market_snapshots_{date_tag}.csv", ms_bytes)

    _write_csv(
        inbox / f"orders_cmd_{date_tag}.csv",
        [
            "schema_version",
            "run_id",
            "cmd_ts",
            "cmd_seq",
            "action",
            "ticker",
            "side",
            "qty",
            "order_type",
            "limit_price",
            "time_in_force",
            "candidate_id",
            "decision_id",
            "client_order_id",
        ],
        [["OC.v1", "R1", "2025-12-24T09:01:00+09:00", "1", "PLACE", "7203", "BUY", "100", "LIMIT", "100", "DAY", "C1", "D1", "O1"]],
    )
    _write_csv(
        outbox / f"execution_events_{date_tag}.csv",
        ["schema_version", "run_id", "event_ts", "event_seq", "cmd_seq", "client_order_id", "broker_order_id", "exec_event"],
        [["EE.v1", "R1", "2025-12-24T09:01:00+09:00", "1", "1", "O1", "", "ACK"]],
    )

    result, _ = run_healthcheck(date_tag=date_tag, base_dir=base)
    assert result.exit_code == 0


def test_healthcheck_market_snapshots_encoding_error_is_critical(tmp_path: Path) -> None:
    base = tmp_path
    date_tag = "20251224"

    outbox = base / "output" / "excel" / "outbox"
    inbox = base / "output" / "excel" / "inbox"

    _write_bytes(outbox / f"market_snapshots_{date_tag}.csv", b"\x00\xff\x00\xff\x00\xff")

    _write_csv(
        inbox / f"orders_cmd_{date_tag}.csv",
        [
            "schema_version",
            "run_id",
            "cmd_ts",
            "cmd_seq",
            "action",
            "ticker",
            "side",
            "qty",
            "order_type",
            "limit_price",
            "time_in_force",
            "candidate_id",
            "decision_id",
            "client_order_id",
        ],
        [["OC.v1", "R1", "2025-12-24T09:01:00+09:00", "1", "PLACE", "7203", "BUY", "100", "LIMIT", "100", "DAY", "C1", "D1", "O1"]],
    )
    _write_csv(
        outbox / f"execution_events_{date_tag}.csv",
        ["schema_version", "run_id", "event_ts", "event_seq", "cmd_seq", "client_order_id", "broker_order_id", "exec_event"],
        [["EE.v1", "R1", "2025-12-24T09:01:00+09:00", "1", "1", "O1", "", "ACK"]],
    )

    result, _ = run_healthcheck(date_tag=date_tag, base_dir=base)
    assert result.exit_code == 2


def test_healthcheck_duplicate_cmd_seq_is_critical(tmp_path: Path) -> None:
    base = tmp_path
    date_tag = "20251224"

    outbox = base / "output" / "excel" / "outbox"
    inbox = base / "output" / "excel" / "inbox"

    _write_csv(
        outbox / f"market_snapshots_{date_tag}.csv",
        ["schema_version", "run_id", "snap_ts", "ticker", "last"],
        [["MS.v1", "R1", "2025-12-24T09:00:00+09:00", "7203", "100"]],
    )

    _write_csv(
        inbox / f"orders_cmd_{date_tag}.csv",
        [
            "schema_version",
            "run_id",
            "cmd_ts",
            "cmd_seq",
            "action",
            "ticker",
            "side",
            "qty",
            "order_type",
            "limit_price",
            "time_in_force",
            "candidate_id",
            "decision_id",
            "client_order_id",
        ],
        [
            ["OC.v1", "R1", "2025-12-24T09:01:00+09:00", "1", "PLACE", "7203", "BUY", "100", "LIMIT", "100", "DAY", "C1", "D1", "O1"],
            ["OC.v1", "R1", "2025-12-24T09:01:00+09:00", "1", "PLACE", "7203", "BUY", "100", "LIMIT", "100", "DAY", "C1", "D1", "O1"],
        ],
    )

    _write_csv(
        outbox / f"execution_events_{date_tag}.csv",
        ["schema_version", "run_id", "event_ts", "event_seq", "cmd_seq", "client_order_id", "broker_order_id", "exec_event"],
        [["EE.v1", "R1", "2025-12-24T09:01:00+09:00", "1", "1", "O1", "", "ACK"]],
    )

    result, _ = run_healthcheck(date_tag=date_tag, base_dir=base)
    assert result.exit_code == 2


def test_healthcheck_duplicate_ack_like_is_critical(tmp_path: Path) -> None:
    base = tmp_path
    date_tag = "20251224"

    outbox = base / "output" / "excel" / "outbox"
    inbox = base / "output" / "excel" / "inbox"

    _write_csv(
        outbox / f"market_snapshots_{date_tag}.csv",
        ["schema_version", "run_id", "snap_ts", "ticker", "last"],
        [["MS.v1", "R1", "2025-12-24T09:00:00+09:00", "7203", "100"]],
    )
    _write_csv(
        inbox / f"orders_cmd_{date_tag}.csv",
        [
            "schema_version",
            "run_id",
            "cmd_ts",
            "cmd_seq",
            "action",
            "ticker",
            "side",
            "qty",
            "order_type",
            "limit_price",
            "time_in_force",
            "candidate_id",
            "decision_id",
            "client_order_id",
        ],
        [["OC.v1", "R1", "2025-12-24T09:01:00+09:00", "1", "PLACE", "7203", "BUY", "100", "LIMIT", "100", "DAY", "C1", "D1", "O1"]],
    )
    _write_csv(
        outbox / f"execution_events_{date_tag}.csv",
        [
            "schema_version",
            "run_id",
            "event_ts",
            "event_seq",
            "cmd_seq",
            "client_order_id",
            "broker_order_id",
            "exec_event",
        ],
        [
            ["EE.v1", "R1", "2025-12-24T09:01:00+09:00", "1", "1", "O1", "", "ACK"],
            ["EE.v1", "R1", "2025-12-24T09:01:01+09:00", "2", "1", "O1", "", "ACK"],
        ],
    )

    result, _ = run_healthcheck(date_tag=date_tag, base_dir=base)
    assert result.exit_code == 2
