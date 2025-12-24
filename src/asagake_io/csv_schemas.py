from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

DT_V1 = "DT.v1"
MS_V1 = "MS.v1"
OC_V1 = "OC.v1"
EE_V1 = "EE.v1"


@dataclass(frozen=True)
class ColumnSpec:
    name: str
    typ: str
    required: bool
    description: str = ""
    enum: Optional[Tuple[str, ...]] = None


def _col(
    name: str,
    typ: str,
    required: bool,
    description: str = "",
    enum: Optional[Sequence[str]] = None,
) -> ColumnSpec:
    enum_t = tuple(enum) if enum is not None else None
    return ColumnSpec(name=name, typ=typ, required=required, description=description, enum=enum_t)


SCHEMAS: Dict[str, List[ColumnSpec]] = {
    DT_V1: [
        _col("schema_version", "str", True, "Schema tag (DT.v1). Must be first."),
        _col("run_id", "str", True, "Run identifier."),
        _col(
            "env",
            "str",
            True,
            "DEMO/LIVE/REPLAY/BT/BRIDGE_SHADOW.",
            enum=["DEMO", "LIVE", "REPLAY", "BT", "BRIDGE_SHADOW"],
        ),
        _col("engine", "str", True, "VBA/PY/HYBRID.", enum=["VBA", "PY", "HYBRID"]),
        _col("engine_version", "str", True, "Git commit/tag."),
        _col("trade_date", "str", True, "Trading date YYYY-MM-DD (JST)."),
        _col("event_ts", "str", True, "Event timestamp (ISO8601)."),
        _col("event_seq", "int", True, "Monotonic within run_id."),
        _col(
            "event_type",
            "str",
            True,
            "Event kind.",
            enum=[
                "MARKET_SNAPSHOT",
                "FEATURES_COMPUTED",
                "FILTER_EVAL",
                "DECISION",
                "ORDER_CMD",
                "ORDER_ACK",
                "FILL",
                "POSITION",
                "EXIT",
                "RISK_STOP",
                "ERROR",
            ],
        ),
        _col("source", "str", True, "EXCEL_RSS / YAHOO_1M / SIM / BROKER, etc."),
        _col("ticker", "str", True, "Ticker."),
        _col("session", "str", False, "Session label."),
        _col("candidate_id", "str", False, "Stable candidate id."),
        _col("decision_id", "str", False, "Decision id."),
        _col("snap_ts", "str", False, "Snapshot timestamp."),
        _col("last", "float", False, "Last price."),
        _col("bid", "float", False, "Best bid."),
        _col("ask", "float", False, "Best ask."),
        _col("bid_size", "float", False, "Bid size."),
        _col("ask_size", "float", False, "Ask size."),
        _col("vwap", "float", False, "VWAP."),
        _col("cum_volume", "float", False, "Cumulative volume."),
        _col("prev_close", "float", False, "Previous close."),
        _col("spread", "float", False, "ask-bid."),
        _col("mid", "float", False, "(bid+ask)/2."),
        _col("nky_last", "float", False, "Nikkei last."),
        _col("topix_last", "float", False, "TOPIX last."),
        _col("nky_ret_bp", "float", False, "Nikkei return (bp)."),
        _col("topix_ret_bp", "float", False, "TOPIX return (bp)."),
        _col("atr_n", "int", False, "ATR period."),
        _col("atr", "float", False, "ATR value."),
        _col("j_raw", "float", False, "Raw J."),
        _col("j_bias_adj", "float", False, "Bias adj."),
        _col("j_gap_adj", "float", False, "Gap adj."),
        _col("j", "float", False, "Final J."),
        _col("j_th", "float", False, "Threshold J_th."),
        _col("signal_mode", "str", False, "j-only / j-cross."),
        _col("j_cross_state", "str", False, "Cross state.", enum=["CROSS_UP", "CROSS_DOWN", "NONE"]),
        _col("allowed", "int", False, "1 allowed, 0 denied."),
        _col("deny_reasons", "str", False, "Deny reasons (semicolon separated)."),
        _col("no_trade_min", "int", False, "NoTrade minutes after open."),
        _col("minutes_from_open", "int", False, "Minutes from open."),
        _col("gap_pct", "float", False, "Gap percentage."),
        _col("gap_ban_pct", "float", False, "Gap ban percentage."),
        _col("trend_driver", "str", False, "NKY / TOPIX."),
        _col("trend_window", "str", False, "Trend window label."),
        _col("trend_bp_th", "float", False, "Trend bp threshold."),
        _col("trend_allowed_policy", "str", False, "ALIGNED_ONLY/BOTH/empty."),
        _col("trend_aligned", "int", False, "1 aligned, 0 mismatch."),
        _col("signal", "str", False, "LONG/SHORT/NONE.", enum=["LONG", "SHORT", "NONE"]),
        _col("action", "str", False, "PLACE/MODIFY/CANCEL/NONE.", enum=["PLACE", "MODIFY", "CANCEL", "NONE"]),
        _col("entry_style", "str", False, "Entry style label."),
        _col("limit_price", "float", False, "Limit price."),
        _col("qty", "int", False, "Quantity."),
        _col("tp_price", "float", False, "TP price."),
        _col("sl_price", "float", False, "SL price."),
        _col("trail_type", "str", False, "Trail type."),
        _col("trail_value", "float", False, "Trail value."),
        _col("tmax_sec", "int", False, "TMAX seconds."),
        _col("client_order_id", "str", False, "Client order id."),
        _col("broker_order_id", "str", False, "Broker order id."),
        _col("order_status", "str", False, "Order status."),
        _col("fill_qty", "int", False, "Fill qty."),
        _col("fill_price", "float", False, "Fill price."),
        _col("pos_qty", "int", False, "Position qty."),
        _col("pos_avg_price", "float", False, "Position avg price."),
        _col("pnl_realized_yen", "float", False, "Realized pnl yen."),
        _col("pnl_unrealized_yen", "float", False, "Unrealized pnl yen."),
        _col("daily_pnl_yen", "float", False, "Daily pnl yen."),
        _col("kill_switch", "int", False, "Kill switch flag."),
        _col("daily_loss_limit_yen", "float", False, "Daily loss limit."),
        _col("consecutive_losses", "int", False, "Consecutive losses."),
        _col("notes", "str", False, "Notes."),
    ],
    MS_V1: [
        _col("schema_version", "str", True, "Schema tag (MS.v1). Must be first."),
        _col("run_id", "str", True, "Run identifier."),
        _col("snap_ts", "str", True, "Snapshot timestamp (ISO8601)."),
        _col("ticker", "str", True, "Ticker."),
        _col("last", "float", False, "Last price."),
        _col("bid", "float", False, "Bid."),
        _col("ask", "float", False, "Ask."),
        _col("vwap", "float", False, "VWAP."),
        _col("cum_volume", "float", False, "Cumulative volume."),
        _col("prev_close", "float", False, "Previous close."),
        _col("bid_size", "float", False, "Bid size."),
        _col("ask_size", "float", False, "Ask size."),
        _col("nky_last", "float", False, "Nikkei last."),
        _col("topix_last", "float", False, "TOPIX last."),
        _col("data_quality", "str", True, "OK/MISSING/STALE.", enum=["OK", "MISSING", "STALE"]),
    ],
    OC_V1: [
        _col("schema_version", "str", True, "Schema tag (OC.v1). Must be first."),
        _col("run_id", "str", True, "Run identifier."),
        _col("cmd_ts", "str", True, "Command timestamp (ISO8601)."),
        _col("cmd_seq", "int", True, "Monotonic within run_id."),
        _col("action", "str", True, "PLACE/MODIFY/CANCEL.", enum=["PLACE", "MODIFY", "CANCEL"]),
        _col("ticker", "str", True, "Ticker."),
        _col("side", "str", True, "BUY/SELL.", enum=["BUY", "SELL"]),
        _col("qty", "int", True, "Quantity."),
        _col("order_type", "str", True, "LIMIT.", enum=["LIMIT"]),
        _col("limit_price", "float", True, "Limit price."),
        _col("time_in_force", "str", False, "TIF."),
        _col("candidate_id", "str", False, "Candidate id."),
        _col("decision_id", "str", False, "Decision id."),
        _col("client_order_id", "str", True, "Client order id."),
        _col("tp_price", "float", False, "TP price."),
        _col("sl_price", "float", False, "SL price."),
        _col("trail_type", "str", False, "Trail type."),
        _col("trail_value", "float", False, "Trail value."),
        _col("tmax_sec", "int", False, "TMAX seconds."),
        _col("reason", "str", False, "Reason code."),
    ],
    EE_V1: [
        _col("schema_version", "str", True, "Schema tag (EE.v1). Must be first."),
        _col("run_id", "str", True, "Run identifier."),
        _col("event_ts", "str", True, "Event timestamp (ISO8601)."),
        _col("event_seq", "int", True, "Monotonic within run_id."),
        _col("cmd_seq", "int", False, "Command seq."),
        _col("client_order_id", "str", True, "Client order id."),
        _col("broker_order_id", "str", False, "Broker order id."),
        _col(
            "exec_event",
            "str",
            True,
            "Execution event kind.",
            enum=["SENT", "ACK", "REJECT", "PARTIAL_FILL", "FILL", "CANCELLED", "EXPIRED"],
        ),
        _col("ticker", "str", True, "Ticker."),
        _col("side", "str", False, "BUY/SELL.", enum=["BUY", "SELL"]),
        _col("qty", "int", False, "Quantity."),
        _col("limit_price", "float", False, "Limit price."),
        _col("fill_qty", "int", False, "Filled qty."),
        _col("fill_price", "float", False, "Filled price."),
        _col("error_code", "str", False, "Error code."),
        _col("error_message", "str", False, "Error message."),
    ],
}


def schema_columns(schema_version: str) -> List[str]:
    if schema_version not in SCHEMAS:
        raise KeyError(f"Unknown schema_version: {schema_version}")
    return [c.name for c in SCHEMAS[schema_version]]


def required_columns(schema_version: str) -> List[str]:
    if schema_version not in SCHEMAS:
        raise KeyError(f"Unknown schema_version: {schema_version}")
    return [c.name for c in SCHEMAS[schema_version] if c.required]


def schema_for_version(schema_version: str) -> List[ColumnSpec]:
    if schema_version not in SCHEMAS:
        raise KeyError(f"Unknown schema_version: {schema_version}")
    return SCHEMAS[schema_version]


def normalize_row_for_columns(row: Dict[str, object], columns: Iterable[str]) -> Dict[str, object]:
    out: Dict[str, object] = {}
    for col in columns:
        val = row.get(col, "")
        out[col] = "" if val is None else val
    return out
