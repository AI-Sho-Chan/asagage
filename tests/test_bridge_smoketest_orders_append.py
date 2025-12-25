from __future__ import annotations

import csv
import subprocess
import sys
from pathlib import Path


def _read_rows(path: Path) -> list[dict[str, str]]:
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        r = csv.DictReader(f)
        return [{k: (v if v is not None else "") for k, v in row.items()} for row in r]


def test_smoketest_orders_cmd_is_append_only(tmp_path: Path) -> None:
    repo_root = Path(__file__).resolve().parents[1]
    tool = repo_root / "tools" / "bridge_smoketest_orders.py"
    assert tool.exists()

    date_tag = "20251224"
    out_dir = tmp_path / "output" / "excel" / "inbox"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"orders_cmd_{date_tag}.csv"

    cmd = [
        sys.executable,
        str(tool),
        "--date",
        date_tag,
        "--run-id",
        "R1",
        "--ticker",
        "7203",
        "--side",
        "BUY",
        "--qty",
        "100",
        "--limit-price",
        "100.0",
    ]

    subprocess.run(cmd, cwd=tmp_path, check=True, capture_output=True, text=True)
    subprocess.run(cmd, cwd=tmp_path, check=True, capture_output=True, text=True)

    rows = _read_rows(out_path)
    assert len(rows) == 2
    assert rows[0].get("cmd_seq") != rows[1].get("cmd_seq")

