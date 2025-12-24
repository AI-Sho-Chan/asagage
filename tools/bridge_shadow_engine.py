from __future__ import annotations

import argparse
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from asagake_core.candidates import CandidateMetadataDefaults, append_candidate_metadata, make_candidate_metadata_defaults
from asagake_io.atomic_writer import atomic_write_csv
from asagake_io.csv_schemas import DT_V1, MS_V1, OC_V1, schema_for_version
from asagake_io.csv_writer import AppendOnlyCsvWriter
from asagake_io.validator import validate_csv


ENV_BRIDGE_SHADOW = "BRIDGE_SHADOW"


@dataclass
class ShadowConfig:
    date_tag: str
    base_dir: Path
    run_id: str
    engine_version: str
    emit_orders: bool
    once: bool
    follow_seconds: int
    poll_interval_sec: float
    max_orders_per_tick: int
    default_qty: int


def _parse_iso8601(value: str) -> Optional[datetime]:
    s = value.strip()
    if not s:
        return None
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00"))
    except ValueError:
        return None


def _iso_now_jst() -> str:
    now = datetime.now(timezone.utc).astimezone(timezone(offset=timezone.utc.utcoffset(None), name="UTC"))
    # Explicit JST offset for consistency with VBA (+09:00).
    # We do not attempt to infer local timezone; callers should treat this as metadata only.
    # Use system local time if possible:
    local = datetime.now().astimezone()
    return local.isoformat(timespec="milliseconds")


def _safe_float(value: object) -> Optional[float]:
    try:
        if value is None:
            return None
        s = str(value).strip()
        if not s:
            return None
        return float(s)
    except ValueError:
        return None


def _safe_int(value: object) -> Optional[int]:
    try:
        if value is None:
            return None
        s = str(value).strip()
        if not s:
            return None
        return int(float(s))
    except ValueError:
        return None


def _read_csv_text(path: Path) -> pd.DataFrame:
    return pd.read_csv(path, dtype=str, keep_default_na=False)


def _get_last_numeric_column_max(path: Path, column: str) -> int:
    if not path.exists():
        return 0
    try:
        df = pd.read_csv(path, usecols=[column], dtype=str, keep_default_na=False)
        if column not in df.columns or df.empty:
            return 0
        seq = pd.to_numeric(df[column], errors="coerce").dropna()
        return int(seq.max()) if len(seq) else 0
    except Exception:
        return 0


def _minutes_from_open_jst(ts: datetime) -> int:
    # Best-effort: use local date and assume Tokyo open at 09:00 local time.
    local = ts.astimezone()
    open_ts = local.replace(hour=9, minute=0, second=0, microsecond=0)
    return int((local - open_ts).total_seconds() // 60)


def _proxy_atr(price: float) -> float:
    # Proxy ATR: 0.10% of price, with a small floor to avoid division by zero.
    return max(0.01, abs(price) * 0.001)


def _compute_j(mid: float, vwap: float) -> Tuple[float, float]:
    atr = _proxy_atr(vwap if vwap else mid)
    j_raw = (mid - vwap) / atr if atr else 0.0
    return j_raw, j_raw


def _signal_from_j(j: float, j_th: float) -> str:
    if j_th <= 0:
        return "NONE"
    if abs(j) < j_th:
        return "NONE"
    return "LONG" if j < 0 else "SHORT"


def _dt_writer(path: Path) -> AppendOnlyCsvWriter:
    cols = [c.name for c in schema_for_version(DT_V1)]
    return AppendOnlyCsvWriter(path=path, columns=cols, schema_version=DT_V1)


def _oc_columns() -> List[str]:
    return [c.name for c in schema_for_version(OC_V1)]


def _build_orders_cmd_row(
    *,
    run_id: str,
    cmd_ts: str,
    cmd_seq: int,
    action: str,
    ticker: str,
    side: str,
    qty: int,
    limit_price: float,
    candidate_id: str,
    decision_id: str,
    client_order_id: str,
    reason: str,
) -> Dict[str, object]:
    return {
        "schema_version": OC_V1,
        "run_id": run_id,
        "cmd_ts": cmd_ts,
        "cmd_seq": cmd_seq,
        "action": action,
        "ticker": ticker,
        "side": side,
        "qty": qty,
        "order_type": "LIMIT",
        "limit_price": f"{limit_price:.6f}",
        "time_in_force": "DAY",
        "candidate_id": candidate_id,
        "decision_id": decision_id,
        "client_order_id": client_order_id,
        "reason": reason,
    }


def run_shadow_engine(
    *,
    config: ShadowConfig,
    candidates_path: Path,
    snapshots_path: Path,
    decision_trace_path: Path,
    orders_cmd_path: Path,
) -> None:
    candidates_df = _read_csv_text(candidates_path)
    base_defaults = make_candidate_metadata_defaults(date_tag=config.date_tag)
    defaults = CandidateMetadataDefaults(
        date_tag=config.date_tag,
        generator_run_id=f"BRIDGE_SHADOW_{config.date_tag}",
        generated_at=base_defaults.generated_at,
        cost_model=base_defaults.cost_model,
        schema_version=base_defaults.schema_version,
    )
    candidates_df = append_candidate_metadata(candidates_df, defaults=defaults)
    candidates_by_ticker: Dict[str, pd.DataFrame] = {
        str(t): df.copy()
        for t, df in candidates_df.groupby("Ticker", dropna=False)
        if str(t).strip()
    }

    dt_writer = _dt_writer(decision_trace_path)
    event_seq = _get_last_numeric_column_max(decision_trace_path, "event_seq")
    cmd_seq = _get_last_numeric_column_max(orders_cmd_path, "cmd_seq")

    print(f"[shadow] candidates={candidates_path} tickers={len(candidates_by_ticker)}")
    print(f"[shadow] snapshots={snapshots_path}")
    print(f"[shadow] decision_trace={decision_trace_path} (event_seq_start={event_seq})")
    if config.emit_orders:
        print(f"[shadow] orders_cmd={orders_cmd_path} (cmd_seq_start={cmd_seq})")

    last_row_count = 0
    start_time = time.time()

    while True:
        if not snapshots_path.exists():
            print("[shadow] snapshots missing; waiting...")
            time.sleep(config.poll_interval_sec)
            continue

        snap_df = _read_csv_text(snapshots_path)
        if len(snap_df) <= last_row_count:
            if config.once:
                break
            if config.follow_seconds > 0 and time.time() - start_time >= config.follow_seconds:
                break
            time.sleep(config.poll_interval_sec)
            continue

        new_df = snap_df.iloc[last_row_count:].copy()
        last_row_count = len(snap_df)

        rows_to_append: List[Dict[str, object]] = []
        new_orders: List[Dict[str, object]] = []
        orders_emitted = 0

        for _, snap in new_df.iterrows():
            ticker = str(snap.get("ticker", "")).strip()
            if not ticker:
                continue
            if ticker not in candidates_by_ticker:
                continue
            snap_ts_s = str(snap.get("snap_ts", "")).strip()
            snap_ts = _parse_iso8601(snap_ts_s) or datetime.now().astimezone()

            last = _safe_float(snap.get("last")) or 0.0
            bid = _safe_float(snap.get("bid"))
            ask = _safe_float(snap.get("ask"))
            vwap = _safe_float(snap.get("vwap")) or 0.0
            prev_close = _safe_float(snap.get("prev_close"))

            if bid is not None and ask is not None and bid > 0 and ask > 0:
                mid = (bid + ask) / 2.0
                spread = ask - bid
            else:
                mid = last
                spread = None

            j_raw, j = _compute_j(mid, vwap) if vwap else (0.0, 0.0)
            minutes_from_open = _minutes_from_open_jst(snap_ts)

            cdf = candidates_by_ticker[ticker]
            for _, cand in cdf.iterrows():
                candidate_id = str(cand.get("candidate_id", "")).strip()
                session = str(cand.get("session", "")).strip()

                j_th = _safe_float(cand.get("J_th")) or _safe_float(cand.get("J_th_base")) or 0.0
                no_trade_min = _safe_int(cand.get("NoTradeMin")) or 0
                gap_ban_pct = _safe_float(cand.get("GapBanPct")) or 0.0

                deny: List[str] = []
                allowed = 1
                if minutes_from_open >= 0 and minutes_from_open < no_trade_min:
                    allowed = 0
                    deny.append("NO_TRADE_MIN")

                gap_pct = None
                if prev_close and prev_close != 0 and vwap:
                    gap_pct = ((mid - prev_close) / prev_close) * 100.0
                    if gap_ban_pct and abs(gap_pct) >= gap_ban_pct:
                        allowed = 0
                        deny.append("GAP_BAN")

                signal = _signal_from_j(j, float(j_th)) if allowed else "NONE"
                action = "PLACE" if signal in {"LONG", "SHORT"} else "NONE"

                decision_id = f"D_{config.date_tag}_{event_seq + 1:06d}"

                def _base_row(event_type: str) -> Dict[str, object]:
                    return {
                        "schema_version": DT_V1,
                        "run_id": config.run_id,
                        "env": ENV_BRIDGE_SHADOW,
                        "engine": "PY",
                        "engine_version": config.engine_version,
                        "trade_date": snap_ts.astimezone().date().isoformat(),
                        "event_ts": snap_ts_s or _iso_now_jst(),
                        "event_seq": None,
                        "event_type": event_type,
                        "source": MS_V1,
                        "ticker": ticker,
                        "session": session,
                        "candidate_id": candidate_id,
                        "decision_id": decision_id,
                        "snap_ts": snap_ts_s or _iso_now_jst(),
                        "last": last,
                        "bid": bid if bid is not None else "",
                        "ask": ask if ask is not None else "",
                        "vwap": vwap if vwap else "",
                        "prev_close": prev_close if prev_close is not None else "",
                        "spread": spread if spread is not None else "",
                        "mid": mid,
                        "atr": _proxy_atr(vwap if vwap else mid),
                        "j_raw": j_raw,
                        "j": j,
                        "j_th": j_th,
                        "no_trade_min": no_trade_min,
                        "minutes_from_open": minutes_from_open,
                        "gap_pct": gap_pct if gap_pct is not None else "",
                        "gap_ban_pct": gap_ban_pct if gap_ban_pct else "",
                        "allowed": allowed,
                        "deny_reasons": ";".join(deny),
                        "signal": signal,
                        "action": action,
                        "limit_price": "",
                        "qty": "",
                    }

                for et in ["MARKET_SNAPSHOT", "FEATURES_COMPUTED", "FILTER_EVAL", "DECISION"]:
                    event_seq += 1
                    row = _base_row(et)
                    row["event_seq"] = event_seq
                    if et == "DECISION" and action == "PLACE":
                        row["entry_style"] = "LIMIT_AHEAD"
                        limit_price = ask if signal == "SHORT" and ask is not None else bid if signal == "LONG" and bid is not None else mid
                        qty = _safe_int(cand.get("OrderQty")) or _safe_int(cand.get("LotSize")) or config.default_qty
                        row["limit_price"] = limit_price
                        row["qty"] = qty
                    rows_to_append.append(row)

                if config.emit_orders and action == "PLACE" and orders_emitted < config.max_orders_per_tick:
                    # Safety guard: emit at most N orders per polling iteration.
                    cmd_seq += 1
                    orders_emitted += 1
                    side = "BUY" if signal == "LONG" else "SELL"
                    limit_price = ask if side == "SELL" and ask is not None else bid if side == "BUY" and bid is not None else mid
                    qty = _safe_int(cand.get("OrderQty")) or _safe_int(cand.get("LotSize")) or config.default_qty
                    client_order_id = f"O_{config.date_tag}_{cmd_seq:06d}"
                    new_orders.append(
                        _build_orders_cmd_row(
                            run_id=config.run_id,
                            cmd_ts=snap_ts_s or _iso_now_jst(),
                            cmd_seq=cmd_seq,
                            action="PLACE",
                            ticker=ticker,
                            side=side,
                            qty=qty,
                            limit_price=float(limit_price),
                            candidate_id=candidate_id,
                            decision_id=decision_id,
                            client_order_id=client_order_id,
                            reason="BRIDGE_SHADOW",
                        )
                    )

        if rows_to_append:
            dt_writer.append_rows(rows_to_append)

        if config.emit_orders and new_orders:
            existing: List[Dict[str, object]] = []
            if orders_cmd_path.exists():
                try:
                    exist_df = pd.read_csv(orders_cmd_path, dtype=str, keep_default_na=False)
                    existing = exist_df.to_dict(orient="records")
                except Exception:
                    existing = []
            all_rows = existing + new_orders
            atomic_write_csv(orders_cmd_path, columns=_oc_columns(), rows=all_rows, encoding="utf-8-sig")

        if config.once:
            break
        if config.follow_seconds > 0 and time.time() - start_time >= config.follow_seconds:
            break

    # Lightweight validation at the end (non-fatal).
    try:
        errs = validate_csv(decision_trace_path, schema_version=DT_V1)
        if errs:
            print(f"[shadow][warn] DT validation errors: {len(errs)}")
    except Exception as exc:
        print(f"[shadow][warn] DT validation failed: {exc!r}")


def main(argv: Optional[List[str]] = None) -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--date", required=True, help="YYYYMMDD")
    p.add_argument("--run-id", required=True, help="Use same run_id as Excel for join.")
    p.add_argument("--engine-version", default="unknown", help="Git commit/tag.")
    p.add_argument("--base-dir", default=str(Path.cwd()), help="Repo base dir (default: CWD)")
    p.add_argument("--candidates", default="output/excel/candidates_nextday.csv")
    p.add_argument("--snapshots", default="", help="market_snapshots_YYYYMMDD.csv (default: outbox path)")
    p.add_argument("--decision-trace", default="", help="decision_trace_YYYYMMDD.csv (default: analysis path)")
    p.add_argument("--emit-orders", action="store_true", help="Actually write orders_cmd (default: off)")
    p.add_argument("--once", action="store_true", help="Process current snapshots and exit")
    p.add_argument("--follow-seconds", type=int, default=0, help="Follow snapshots for N seconds")
    p.add_argument("--poll-interval-sec", type=float, default=1.0)
    p.add_argument("--max-orders-per-tick", type=int, default=1)
    p.add_argument("--default-qty", type=int, default=100)
    args = p.parse_args(argv)

    base_dir = Path(args.base_dir).expanduser().resolve()
    date_tag = args.date

    candidates_path = (base_dir / args.candidates).resolve()
    snapshots_path = (
        Path(args.snapshots).expanduser().resolve()
        if args.snapshots
        else (base_dir / "output" / "excel" / "outbox" / f"market_snapshots_{date_tag}.csv").resolve()
    )
    decision_trace_path = (
        Path(args.decision_trace).expanduser().resolve()
        if args.decision_trace
        else (base_dir / "analysis" / f"decision_trace_{date_tag}.csv").resolve()
    )
    orders_cmd_path = (base_dir / "output" / "excel" / "inbox" / f"orders_cmd_{date_tag}.csv").resolve()

    cfg = ShadowConfig(
        date_tag=date_tag,
        base_dir=base_dir,
        run_id=args.run_id,
        engine_version=args.engine_version,
        emit_orders=bool(args.emit_orders),
        once=bool(args.once),
        follow_seconds=int(args.follow_seconds),
        poll_interval_sec=float(args.poll_interval_sec),
        max_orders_per_tick=int(args.max_orders_per_tick),
        default_qty=int(args.default_qty),
    )

    if not candidates_path.exists():
        print(f"[shadow][error] candidates file missing: {candidates_path}")
        return 2

    run_shadow_engine(
        config=cfg,
        candidates_path=candidates_path,
        snapshots_path=snapshots_path,
        decision_trace_path=decision_trace_path,
        orders_cmd_path=orders_cmd_path,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
