from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path
from typing import Optional


@dataclass(frozen=True)
class ExecEventsSummary:
    run_id: str
    max_cmd_seq: int
    max_event_seq: int
    ack_like_duplicate_cmd_seq: int


def summarize_execution_events(path: Path) -> ExecEventsSummary:
    if not path.exists():
        return ExecEventsSummary(run_id="", max_cmd_seq=0, max_event_seq=0, ack_like_duplicate_cmd_seq=0)

    run_id = ""
    max_cmd = 0
    max_evt = 0
    seen_ack_cmd: set[int] = set()
    dup_ack = 0

    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f)
        try:
            header = next(r)
        except StopIteration:
            return ExecEventsSummary(run_id="", max_cmd_seq=0, max_event_seq=0, ack_like_duplicate_cmd_seq=0)

        idx = {name: i for i, name in enumerate(header)}
        cmd_i = idx.get("cmd_seq")
        evt_i = idx.get("event_seq")
        run_i = idx.get("run_id")
        ev_i = idx.get("exec_event")

        for row in r:
            if not row:
                continue
            if run_i is not None and run_i < len(row) and not run_id:
                run_id = str(row[run_i]).strip()
            if cmd_i is not None and cmd_i < len(row):
                try:
                    cmd = int(float(str(row[cmd_i]).strip() or "0"))
                    if cmd > max_cmd:
                        max_cmd = cmd
                except ValueError:
                    pass
            if evt_i is not None and evt_i < len(row):
                try:
                    evt = int(float(str(row[evt_i]).strip() or "0"))
                    if evt > max_evt:
                        max_evt = evt
                except ValueError:
                    pass

            if ev_i is not None and cmd_i is not None and ev_i < len(row) and cmd_i < len(row):
                ev = str(row[ev_i]).strip().upper()
                if ev in {"ACK", "REJECT"}:
                    try:
                        cmd = int(float(str(row[cmd_i]).strip() or "0"))
                        if cmd in seen_ack_cmd:
                            dup_ack += 1
                        else:
                            seen_ack_cmd.add(cmd)
                    except ValueError:
                        pass

    return ExecEventsSummary(run_id=run_id, max_cmd_seq=max_cmd, max_event_seq=max_evt, ack_like_duplicate_cmd_seq=dup_ack)

