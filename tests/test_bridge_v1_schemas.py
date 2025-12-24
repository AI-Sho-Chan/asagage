from __future__ import annotations

import json
from pathlib import Path

import pytest

from asagake_io.csv_schemas import DT_V1, EE_V1, MS_V1, OC_V1, schema_columns
from asagake_io.csv_writer import make_append_only_writer
from asagake_io.validator import validate_csv


def _load_schema(path: Path) -> dict:
    # Note: the repo stores these .yaml files as JSON (YAML is a superset of JSON),
    # so we can validate without introducing a YAML dependency.
    return json.loads(path.read_text(encoding="utf-8"))


@pytest.mark.parametrize(
    ("schema_version", "schema_path"),
    [
        (DT_V1, Path("schemas/decision_trace_dt_v1.yaml")),
        (MS_V1, Path("schemas/market_snapshot_ms_v1.yaml")),
        (OC_V1, Path("schemas/orders_cmd_oc_v1.yaml")),
        (EE_V1, Path("schemas/execution_events_ee_v1.yaml")),
    ],
)
def test_schema_files_match_python_columns(schema_version: str, schema_path: Path) -> None:
    assert schema_path.exists()
    data = _load_schema(schema_path)
    assert data["schema_version"] == schema_version
    cols = [c["name"] for c in data["columns"]]
    assert cols == schema_columns(schema_version)
    assert cols[0] == "schema_version"


def test_append_writer_creates_utf8sig_header(tmp_path: Path) -> None:
    out = tmp_path / "decision_trace_20251224.csv"
    cols = schema_columns(DT_V1)
    w = make_append_only_writer(out, schema_version=DT_V1, columns=cols)
    w.append_rows(
        [
            {
                "schema_version": DT_V1,
                "run_id": "20251224_REPLAY_TEST_000000",
                "env": "REPLAY",
                "engine": "PY",
                "engine_version": "test",
                "trade_date": "2025-12-24",
                "event_ts": "2025-12-24T09:00:00+09:00",
                "event_seq": 1,
                "event_type": "ERROR",
                "source": "SIM",
                "ticker": "7203",
                "notes": "hello",
            }
        ]
    )
    # second append must not add another header
    w.append_rows(
        [
            {
                "schema_version": DT_V1,
                "run_id": "20251224_REPLAY_TEST_000000",
                "env": "REPLAY",
                "engine": "PY",
                "engine_version": "test",
                "trade_date": "2025-12-24",
                "event_ts": "2025-12-24T09:00:01+09:00",
                "event_seq": 2,
                "event_type": "ERROR",
                "source": "SIM",
                "ticker": "7203",
                "notes": "hello2",
            }
        ]
    )
    text = out.read_text(encoding="utf-8-sig")
    assert text.count("schema_version,") == 1
    errs = validate_csv(out, schema_version=DT_V1)
    assert not errs

