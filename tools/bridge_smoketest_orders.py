from __future__ import annotations

import argparse
import csv
import sys
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional, Tuple

SRC = (Path(__file__).resolve().parents[1] / "src").resolve()
if SRC.exists():
    sys.path.insert(0, str(SRC))

from asagake_io.atomic_writer import atomic_write_csv  # noqa: E402
from asagake_io.csv_schemas import OC_V1, schema_columns  # noqa: E402


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Bridge.v1 smoketest: write a single OrdersCmd (OC.v1) for Excel DEMO to consume."
    )
    ap.add_argument("--date", required=True, help="YYYYMMDD (JST) for orders_cmd_YYYYMMDD.csv")
    ap.add_argument("--run-id", required=True, help="run_id to embed (e.g. 20251224_DEMO_PC01_001)")
    ap.add_argument("--ticker", required=True, help="Ticker code (e.g. 7203 or 7203.T)")
    ap.add_argument("--side", required=True, choices=["BUY", "SELL"])
    ap.add_argument("--qty", required=True, type=int)
    ap.add_argument("--limit-price", required=True, type=float)
    ap.add_argument("--cmd-seq-start", type=int, default=None, help="Optional starting cmd_seq override.")
    ap.add_argument("--action", default="PLACE", choices=["PLACE", "MODIFY", "CANCEL"])
    ap.add_argument("--candidate-id", default="", help="Optional candidate_id (empty is OK for smoketest).")
    ap.add_argument("--decision-id", default="", help="Optional decision_id (empty is OK for smoketest).")
    return ap.parse_args()


def _read_existing_max_cmd_seq(path: Path) -> int:
    if not path.exists():
        return 0
    try:
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            r = csv.DictReader(f)
            max_seq = 0
            for row in r:
                raw = (row.get("cmd_seq") or "").strip()
                if not raw:
                    continue
                try:
                    seq = int(float(raw))
                except ValueError:
                    continue
                max_seq = max(max_seq, seq)
            return max_seq
    except OSError:
        return 0


def _load_existing_orders_cmd(path: Path) -> Tuple[list[str], list[dict[str, str]]]:
    if not path.exists():
        return [], []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        columns = list(r.fieldnames or [])
        rows: list[dict[str, str]] = []
        for row in r:
            if not row:
                continue
            rows.append({k: (v if v is not None else "") for k, v in row.items()})
        return columns, rows


def _ensure_columns(existing: list[str], required: list[str]) -> list[str]:
    out = list(existing)
    seen = {c: True for c in out}
    for col in required:
        if col not in seen:
            out.append(col)
            seen[col] = True
    return out


def _next_unused_cmd_seq(existing_rows: Iterable[dict[str, str]], start: int) -> int:
    used: set[int] = set()
    max_seq = 0
    for row in existing_rows:
        raw = (row.get("cmd_seq") or "").strip()
        if not raw:
            continue
        try:
            seq = int(float(raw))
        except ValueError:
            continue
        used.add(seq)
        if seq > max_seq:
            max_seq = seq

    seq = max(max_seq + 1, start) if start > 0 else (max_seq + 1)
    while seq in used:
        seq += 1
    return seq


def main() -> None:
    args = parse_args()

    date_tag = args.date.strip()
    inbox = Path("output/excel/inbox")
    out = inbox / f"orders_cmd_{date_tag}.csv"

    schema_cols = schema_columns(OC_V1)
    existing_cols, existing_rows = _load_existing_orders_cmd(out)
    cols = _ensure_columns(existing_cols if existing_cols else schema_cols, schema_cols)

    start = args.cmd_seq_start or 0
    cmd_seq = _next_unused_cmd_seq(existing_rows, start)

    now = datetime.now().astimezone().isoformat(timespec="seconds")
    client_order_id = f"O_{args.date}_{cmd_seq:06d}"

    row = {
        "schema_version": OC_V1,
        "run_id": args.run_id,
        "cmd_ts": now,
        "cmd_seq": cmd_seq,
        "action": args.action,
        "ticker": args.ticker,
        "side": args.side,
        "qty": int(args.qty),
        "order_type": "LIMIT",
        "limit_price": float(args.limit_price),
        "time_in_force": "DAY",
        "candidate_id": args.candidate_id,
        "decision_id": args.decision_id,
        "client_order_id": client_order_id,
        "reason": "BRIDGE_SMOKETEST",
    }

    all_rows: list[dict[str, object]] = []
    for existing in existing_rows:
        out_row: dict[str, object] = {}
        for col in cols:
            out_row[col] = existing.get(col, "")
        all_rows.append(out_row)

    new_row: dict[str, object] = {}
    for col in cols:
        new_row[col] = row.get(col, "")
    all_rows.append(new_row)

    atomic_write_csv(out, columns=cols, rows=all_rows)
    print(
        {
            "written": str(out),
            "schema_version": OC_V1,
            "cmd_seq": cmd_seq,
            "client_order_id": client_order_id,
        }
    )


if __name__ == "__main__":
    main()
