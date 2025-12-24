from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import Optional

import pandas as pd

from .ids import candidate_id_from_row, params_hash_from_row, safe_generator_run_id


CAND_V1_SCHEMA_VERSION = "CAND.v1"
DEFAULT_COST_MODEL = "CM_v1"


def now_iso_jst() -> str:
    return (
        datetime.now(timezone.utc)
        .astimezone(timezone(timedelta(hours=9)))
        .isoformat(timespec="seconds")
    )


@dataclass(frozen=True)
class CandidateMetadataDefaults:
    date_tag: str
    generator_run_id: str
    generated_at: str
    cost_model: str = DEFAULT_COST_MODEL
    schema_version: str = CAND_V1_SCHEMA_VERSION


def make_candidate_metadata_defaults(*, date_tag: str, git_short_sha: Optional[str] = None) -> CandidateMetadataDefaults:
    date_tag2 = date_tag or "unknown"
    return CandidateMetadataDefaults(
        date_tag=date_tag2,
        generator_run_id=safe_generator_run_id(date_tag=date_tag2, git_short_sha=git_short_sha),
        generated_at=now_iso_jst(),
    )


def append_candidate_metadata(
    df: pd.DataFrame,
    *,
    defaults: CandidateMetadataDefaults,
) -> pd.DataFrame:
    if df.empty:
        return df

    # Append-only: never rename/reorder existing columns. Add missing columns to the right.
    if "schema_version" not in df.columns:
        df["schema_version"] = defaults.schema_version
    if "candidate_id" not in df.columns:
        df["candidate_id"] = df.apply(candidate_id_from_row, axis=1)
    if "params_hash" not in df.columns:
        df["params_hash"] = df.apply(params_hash_from_row, axis=1)
    if "generated_at" not in df.columns:
        df["generated_at"] = defaults.generated_at
    if "generator_run_id" not in df.columns:
        df["generator_run_id"] = defaults.generator_run_id
    if "cost_model" not in df.columns:
        df["cost_model"] = defaults.cost_model
    return df

