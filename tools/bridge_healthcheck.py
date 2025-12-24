from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from datetime import datetime
from io import StringIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd


@dataclass(frozen=True)
class HealthResult:
    ok: bool
    exit_code: int
    csv_path: Path
    txt_path: Path


def _parse_iso8601_maybe(value: object) -> Optional[datetime]:
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00"))
    except ValueError:
        return None


def _safe_int_series(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(-1).astype("int64")


def _decode_csv_bytes_with_recovery(raw: bytes) -> Tuple[str, List[str]]:
    """
    Decode CSV bytes robustly for Excel/VBA-produced files.

    Historical VBA code sometimes wrote Variant(Byte()) directly to Put #,
    which prepends a Variant header and moves the UTF-8 BOM away from byte 0.
    This helper trims such prefixes (BOM/header not at start) so the CSV remains readable.
    """

    warnings: List[str] = []
    bom = b"\xEF\xBB\xBF"

    bom_idx = raw.find(bom)
    if bom_idx > 0:
        warnings.append(f"prefix_bom_trimmed_at={bom_idx}")
        raw = raw[bom_idx:]
    elif bom_idx == -1:
        header_idx = raw.find(b"schema_version,")
        if header_idx > 0:
            warnings.append(f"prefix_header_trimmed_at={header_idx}")
            raw = raw[header_idx:]

    try:
        return raw.decode("utf-8-sig"), warnings
    except UnicodeDecodeError as exc:
        warnings.append(f"encoding_error={exc!r}")
        return "", warnings


def _load_csv(path: Path) -> Tuple[Optional[pd.DataFrame], List[str]]:
    """
    Load a UTF-8-sig CSV with recovery.

    Returns (df, warnings). df=None means "could not be parsed" and callers must not crash.
    """

    try:
        return pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig"), []
    except UnicodeDecodeError:
        text, warnings = _decode_csv_bytes_with_recovery(path.read_bytes())
        if not text:
            return None, warnings
        try:
            return pd.read_csv(StringIO(text), dtype=str, keep_default_na=False), warnings
        except Exception as exc:  # pragma: no cover - defensive
            warnings.append(f"parse_error={exc!r}")
            return None, warnings
    except Exception as exc:  # pragma: no cover - defensive
        return None, [f"parse_error={exc!r}"]


def _metrics_to_df(date_tag: str, metrics: List[Dict[str, object]]) -> pd.DataFrame:
    rows: List[Dict[str, object]] = []
    for metric in metrics:
        row = {"date_tag": date_tag, **metric}
        rows.append(row)
    return pd.DataFrame(rows)


def run_healthcheck(*, date_tag: str, base_dir: Path) -> Tuple[HealthResult, List[str]]:
    analysis_dir = base_dir / "analysis"
    analysis_dir.mkdir(parents=True, exist_ok=True)

    inbox_dir = base_dir / "output" / "excel" / "inbox"
    outbox_dir = base_dir / "output" / "excel" / "outbox"

    ms_path = outbox_dir / f"market_snapshots_{date_tag}.csv"
    oc_path = inbox_dir / f"orders_cmd_{date_tag}.csv"
    ee_path = outbox_dir / f"execution_events_{date_tag}.csv"

    csv_out = analysis_dir / f"bridge_health_{date_tag}.csv"
    txt_out = analysis_dir / f"bridge_health_{date_tag}.txt"

    msgs: List[str] = []
    metrics: List[Dict[str, object]] = []
    critical = False

    def add(check: str, metric: str, value: object, severity: str = "INFO", details: str = "") -> None:
        nonlocal critical
        metrics.append(
            {
                "check": check,
                "metric": metric,
                "value": value,
                "severity": severity,
                "details": details,
            }
        )
        if severity == "CRITICAL":
            critical = True

    def missing_file(path: Path, label: str) -> None:
        add("files", f"{label}_exists", False, "CRITICAL", f"missing: {path}")
        msgs.append(f"[CRITICAL] {label} missing: {path}")

    if not ms_path.exists():
        missing_file(ms_path, "market_snapshots")
        ms_df = None
    else:
        add("files", "market_snapshots_exists", True)
        ms_df, ms_warn = _load_csv(ms_path)
        if ms_df is None:
            add("market_snapshots", "encoding_error", True, "CRITICAL", "failed to parse market_snapshots")
            add("market_snapshots", "read_warnings", "|".join(ms_warn), "CRITICAL")
            msgs.append("[CRITICAL] market_snapshots could not be parsed (encoding/csv error)")
        else:
            if ms_warn:
                add("market_snapshots", "read_warnings", "|".join(ms_warn), "WARN")
                msgs.append(f"[WARN] market_snapshots recovered with warnings: {' | '.join(ms_warn)}")

            add("market_snapshots", "rows", len(ms_df))
            if "snap_ts" in ms_df.columns:
                last_ts = _parse_iso8601_maybe(ms_df["snap_ts"].iloc[-1]) if len(ms_df) else None
                add("market_snapshots", "last_snap_ts", last_ts.isoformat() if last_ts else "")
            else:
                add("market_snapshots", "missing_column_snap_ts", True, "CRITICAL")
                msgs.append("[CRITICAL] market_snapshots missing required column snap_ts")

            for col in ["schema_version", "run_id", "snap_ts", "ticker"]:
                if col not in ms_df.columns:
                    add("market_snapshots", f"missing_required_{col}", True, "CRITICAL")

            if all(c in ms_df.columns for c in ["run_id", "snap_ts", "ticker"]):
                dup = ms_df.duplicated(subset=["run_id", "snap_ts", "ticker"]).sum()
                add("market_snapshots", "duplicate_rows", int(dup), "WARN" if dup else "INFO")

            if "data_quality" in ms_df.columns:
                bad = (ms_df["data_quality"].astype(str).str.upper() != "OK").sum()
                add(
                    "market_snapshots",
                    "non_ok_quality_rows",
                    int(bad),
                    "WARN" if bad else "INFO",
                )

    if not oc_path.exists():
        missing_file(oc_path, "orders_cmd")
        oc_df = None
    else:
        add("files", "orders_cmd_exists", True)
        oc_df, oc_warn = _load_csv(oc_path)
        if oc_df is None:
            add("orders_cmd", "encoding_error", True, "CRITICAL", "failed to parse orders_cmd")
            add("orders_cmd", "read_warnings", "|".join(oc_warn), "CRITICAL")
            msgs.append("[CRITICAL] orders_cmd could not be parsed (encoding/csv error)")
        else:
            if oc_warn:
                add("orders_cmd", "read_warnings", "|".join(oc_warn), "WARN")

            add("orders_cmd", "rows", len(oc_df))
            if "cmd_seq" not in oc_df.columns:
                add("orders_cmd", "missing_column_cmd_seq", True, "CRITICAL")
                msgs.append("[CRITICAL] orders_cmd missing required column cmd_seq")
            else:
                seq = _safe_int_series(oc_df["cmd_seq"])
                dup = seq.duplicated().sum()
                add("orders_cmd", "duplicate_cmd_seq", int(dup), "CRITICAL" if dup else "INFO")
                uniq = seq.drop_duplicates().sort_values()
                gaps = int((uniq.diff() > 1).sum()) if len(uniq) else 0
                add("orders_cmd", "cmd_seq_gap_count", gaps, "WARN" if gaps else "INFO")
                if len(uniq):
                    add("orders_cmd", "cmd_seq_min", int(uniq.iloc[0]))
                    add("orders_cmd", "cmd_seq_max", int(uniq.iloc[-1]))

    if not ee_path.exists():
        missing_file(ee_path, "execution_events")
        ee_df = None
    else:
        add("files", "execution_events_exists", True)
        ee_df, ee_warn = _load_csv(ee_path)
        if ee_df is None:
            add("execution_events", "encoding_error", True, "CRITICAL", "failed to parse execution_events")
            add("execution_events", "read_warnings", "|".join(ee_warn), "CRITICAL")
            msgs.append("[CRITICAL] execution_events could not be parsed (encoding/csv error)")
        else:
            if ee_warn:
                add("execution_events", "read_warnings", "|".join(ee_warn), "WARN")

            add("execution_events", "rows", len(ee_df))
            for col in ["event_seq", "cmd_seq", "exec_event"]:
                if col not in ee_df.columns:
                    add("execution_events", f"missing_required_{col}", True, "CRITICAL")

            if "cmd_seq" in ee_df.columns:
                ee_seq = _safe_int_series(ee_df["cmd_seq"])
                dup = (
                    ee_df.duplicated(subset=["cmd_seq", "exec_event", "client_order_id"]).sum()
                    if "client_order_id" in ee_df.columns
                    else 0
                )
                add(
                    "execution_events",
                    "duplicate_cmd_seq_events",
                    int(dup),
                    "WARN" if dup else "INFO",
                )
                if "exec_event" in ee_df.columns:
                    ack_like = ee_df["exec_event"].astype(str).str.upper().isin(["ACK", "REJECT"])
                    if ack_like.any():
                        ack_dups = ee_df[ack_like].duplicated(subset=["cmd_seq"]).sum()
                        add(
                            "execution_events",
                            "ack_like_duplicate_cmd_seq",
                            int(ack_dups),
                            "CRITICAL" if ack_dups else "INFO",
                            "Excel may have reprocessed commands after restart if this is >0.",
                        )

    if (
        oc_df is not None
        and ee_df is not None
        and "cmd_seq" in oc_df.columns
        and "cmd_seq" in ee_df.columns
        and "client_order_id" in oc_df.columns
        and "client_order_id" in ee_df.columns
    ):
        oc_pairs = set(
            zip(
                _safe_int_series(oc_df["cmd_seq"]).tolist(),
                oc_df["client_order_id"].astype(str).tolist(),
            )
        )
        ee_pairs = set(
            zip(
                _safe_int_series(ee_df["cmd_seq"]).tolist(),
                ee_df["client_order_id"].astype(str).tolist(),
            )
        )
        oc_pairs = {(s, o) for (s, o) in oc_pairs if s >= 0 and o.strip()}
        ee_pairs = {(s, o) for (s, o) in ee_pairs if s >= 0 and o.strip()}

        pending = sorted(oc_pairs - ee_pairs)
        add("reconcile", "orders_without_events", len(pending), "WARN" if pending else "INFO")
        if pending:
            add(
                "reconcile",
                "orders_without_events_sample",
                ";".join(f"{s}:{o}" for s, o in pending[:10]),
                "INFO",
            )

        unknown = sorted(ee_pairs - oc_pairs)
        add("reconcile", "events_without_orders", len(unknown), "WARN" if unknown else "INFO")
        if unknown:
            add(
                "reconcile",
                "events_without_orders_sample",
                ";".join(f"{s}:{o}" for s, o in unknown[:10]),
                "INFO",
            )

        if "exec_event" in ee_df.columns:
            ev = ee_df["exec_event"].astype(str).str.upper()
            reject_rate = (ev == "REJECT").mean() if len(ev) else 0.0
            add("reconcile", "reject_rate", float(reject_rate), "WARN" if reject_rate >= 0.05 else "INFO")

    out_df = _metrics_to_df(date_tag, metrics)
    out_df.to_csv(csv_out, index=False, encoding="utf-8-sig")

    summary_lines = [
        f"ASAGAKE Bridge healthcheck ({date_tag})",
        f"base_dir={base_dir}",
        f"market_snapshots={ms_path}",
        f"orders_cmd={oc_path}",
        f"execution_events={ee_path}",
        "",
    ]
    for m in metrics:
        if m["severity"] in {"CRITICAL", "WARN"}:
            summary_lines.append(
                f"[{m['severity']}] {m['check']}.{m['metric']}={m['value']}"
                + (f" ({m['details']})" if m.get("details") else "")
            )
    if len(summary_lines) == 6:
        summary_lines.append("OK: no warnings.")
    txt_out.write_text("\n".join(summary_lines) + "\n", encoding="utf-8-sig")

    exit_code = 2 if critical else 0
    result = HealthResult(ok=not critical, exit_code=exit_code, csv_path=csv_out, txt_path=txt_out)
    return result, msgs


def main(argv: Optional[List[str]] = None) -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--date", required=True, help="YYYYMMDD")
    p.add_argument("--base-dir", default=str(Path.cwd()), help="Repo base dir (default: CWD)")
    args = p.parse_args(argv)

    base_dir = Path(args.base_dir).expanduser().resolve()
    result, _ = run_healthcheck(date_tag=args.date, base_dir=base_dir)
    print(f"Wrote: {result.csv_path}")
    print(f"Wrote: {result.txt_path}")
    return result.exit_code


if __name__ == "__main__":
    raise SystemExit(main())
