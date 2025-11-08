from __future__ import annotations

from pathlib import Path
import sys
import subprocess


WB = Path("C:/AI/asagake/SHINSOKU.xlsm")
ROOT = Path(__file__).resolve().parent.parent


def run_py(mod: str, *args: str) -> None:
    cmd = [sys.executable, str(ROOT / mod), *map(str, args)]
    print("[run]", " ".join(cmd))
    proc = subprocess.run(cmd)
    if proc.returncode != 0:
        raise SystemExit(proc.returncode)


def main() -> None:
    # 1) Backup workbook
    run_py("scripts/backup_workbook.py", str(WB))

    # 2) Install macros module(s) without opening UI
    run_py("scripts/excel_install_macros.py", str(WB), str(ROOT / "AutoTrader.bas"))

    # 3) Ensure ExecMon sheet for executions (non-fatal if RSS not present)
    run_py("scripts/setup_execmon.py")

    # 4) Burn dashboard realtime formulas (RSS + DynamicQty)
    run_py("scripts/burn_realtime_formulas.py")

    # 5) Cleanup dashboard spares (right of PlanTag)
    run_py("scripts/cleanup_dashboard.py")

    print("Maintenance completed.")


if __name__ == "__main__":
    main()

