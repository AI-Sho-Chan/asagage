from __future__ import annotations

import pandas as pd

from asagake_core.candidates import CandidateMetadataDefaults, append_candidate_metadata
from asagake_core.ids import candidate_id_from_row, params_hash_from_row


def test_candidate_id_is_deterministic() -> None:
    row = {
        "Ticker": "7203.T",
        "session": "AM0930",
        "SignalMode": "j-only",
        "J_th": 1.2,
        "TPk": 1.0,
        "SLk": 2.0,
        "TMAX": 180,
        "ATR_n": 14,
        "BudgetFactor_row": 1.0,
        "NoTradeMin": 5,
        "GapBanPct": 3.0,
        "trend_driver": "NKY",
        "trend_window": "30",
        "trend_bp_th": 12,
        "trend_allowed_policy": "ALIGNED_ONLY",
    }

    cid1 = candidate_id_from_row(row)
    cid2 = candidate_id_from_row(row)
    assert cid1 == cid2

    h1 = params_hash_from_row(row)
    h2 = params_hash_from_row(row)
    assert h1 == h2


def test_append_candidate_metadata_is_right_append_only() -> None:
    df = pd.DataFrame(
        [
            {
                "Ticker": "7203.T",
                "session": "AM0930",
                "SignalMode": "j-only",
                "J_th": 1.2,
                "TPk": 1.0,
                "SLk": 2.0,
                "TMAX": 180,
                "ATR_n": 14,
                "BudgetFactor_row": 1.0,
                "NoTradeMin": 5,
                "GapBanPct": 3.0,
            }
        ]
    )
    original_cols = list(df.columns)

    defaults = CandidateMetadataDefaults(
        date_tag="20251224",
        generator_run_id="AGG_20251224_test",
        generated_at="2025-12-24T09:00:00+09:00",
        cost_model="CM_v1",
        schema_version="CAND.v1",
    )
    out = append_candidate_metadata(df.copy(), defaults=defaults)

    assert out.columns[: len(original_cols)].tolist() == original_cols
    assert out.columns[-6:].tolist() == [
        "schema_version",
        "candidate_id",
        "params_hash",
        "generated_at",
        "generator_run_id",
        "cost_model",
    ]
    assert out.loc[0, "schema_version"] == "CAND.v1"
    assert out.loc[0, "generator_run_id"] == "AGG_20251224_test"
    assert out.loc[0, "generated_at"] == "2025-12-24T09:00:00+09:00"
    assert isinstance(out.loc[0, "candidate_id"], str) and out.loc[0, "candidate_id"].startswith("C_")
    assert isinstance(out.loc[0, "params_hash"], str) and len(out.loc[0, "params_hash"]) == 40

