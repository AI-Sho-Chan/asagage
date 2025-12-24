from __future__ import annotations

import hashlib
from dataclasses import dataclass
from typing import Mapping, Optional


def _norm_str(value: object) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    if s.lower() in {"nan", "none"}:
        return ""
    return s


def _norm_num(value: object) -> str:
    s = _norm_str(value)
    if s == "":
        return ""
    try:
        f = float(s)
    except ValueError:
        return s
    if f != f:  # NaN
        return ""
    # Stable, locale-independent format (no commas, dot decimal).
    txt = f"{f:.10f}".rstrip("0").rstrip(".")
    return txt if txt != "-0" else "0"


@dataclass(frozen=True)
class CandidateIdParts:
    ticker: str
    session: str
    signal_mode: str
    j_th: str
    tpk: str
    slk: str
    tmax: str
    atr_n: str
    budget_factor: str
    no_trade_min: str
    gap_ban_pct: str
    trend_driver: str
    trend_window: str
    trend_bp_th: str
    trend_allowed_policy: str


def candidate_id_parts_from_row(row: Mapping[str, object]) -> CandidateIdParts:
    return CandidateIdParts(
        ticker=_norm_str(row.get("Ticker") or row.get("ticker") or row.get("code")),
        session=_norm_str(row.get("session") or row.get("Session") or row.get("plan_tag")),
        signal_mode=_norm_str(row.get("SignalMode") or row.get("signal_mode")),
        j_th=_norm_num(row.get("J_th") or row.get("J_th_base") or row.get("j_th")),
        tpk=_norm_num(row.get("TPk") or row.get("tpk")),
        slk=_norm_num(row.get("SLk") or row.get("slk")),
        tmax=_norm_num(row.get("TMAX") or row.get("tmax") or row.get("tmax_sec")),
        atr_n=_norm_num(row.get("ATR_n") or row.get("atr_n")),
        budget_factor=_norm_num(row.get("BudgetFactor_row") or row.get("budget_factor")),
        no_trade_min=_norm_num(row.get("NoTradeMin") or row.get("no_trade_min")),
        gap_ban_pct=_norm_num(row.get("GapBanPct") or row.get("gap_ban_pct") or row.get("gapban_pct")),
        trend_driver=_norm_str(row.get("trend_driver") or row.get("TrendDriver")),
        trend_window=_norm_str(row.get("trend_window") or row.get("TrendWindow")),
        trend_bp_th=_norm_num(row.get("trend_bp_th") or row.get("TrendBpTh")),
        trend_allowed_policy=_norm_str(
            row.get("trend_allowed_policy") or row.get("TrendAllowedPolicy")
        ),
    )


def params_hash_from_parts(parts: CandidateIdParts) -> str:
    payload = "\x1f".join(
        [
            parts.ticker,
            parts.session,
            parts.signal_mode,
            parts.j_th,
            parts.tpk,
            parts.slk,
            parts.tmax,
            parts.atr_n,
            parts.budget_factor,
            parts.no_trade_min,
            parts.gap_ban_pct,
            parts.trend_driver,
            parts.trend_window,
            parts.trend_bp_th,
            parts.trend_allowed_policy,
        ]
    )
    return hashlib.sha1(payload.encode("utf-8")).hexdigest()


def params_hash_from_row(row: Mapping[str, object]) -> str:
    return params_hash_from_parts(candidate_id_parts_from_row(row))


def candidate_id_from_row(row: Mapping[str, object], *, hash_len: int = 8) -> str:
    parts = candidate_id_parts_from_row(row)
    h = params_hash_from_parts(parts)[:hash_len]
    ticker = parts.ticker or "UNKNOWN"
    session = parts.session or "NA"
    return f"C_{ticker}_{session}_{h}"


def safe_generator_run_id(*, date_tag: str, git_short_sha: Optional[str] = None) -> str:
    suffix = git_short_sha or "nogit"
    return f"AGG_{date_tag}_{suffix}"

