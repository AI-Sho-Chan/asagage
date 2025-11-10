#!/usr/bin/env python3
"""
ASAGAKE workbook auto-repair helper (headerless VBA injection).

Steps:
1. Create a timestamped backup of the target workbook.
2. Rebuild NewDashboardV2 layout (repair_asagake_dashboard -> restore formulas).
3. Reinstall AutoTraderAdvanced/cDashboardWatcher via AddFromString (no Attribute/VERSION headers).
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import shutil
import subprocess
import sys
from pathlib import Path

try:
    import win32com.client as win32  # type: ignore
except Exception:  # pragma: no cover
    win32 = None

REPO_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_WORKBOOK = REPO_ROOT / "ASAGAKE.xlsm"
AUTO_SRC = REPO_ROOT / "excel" / "AutoTraderAdvanced.bas"


def run(cmd: list[str]) -> None:
    subprocess.run([sys.executable, *cmd], cwd=REPO_ROOT, check=True)


def _sanitize_module_text(raw: str) -> str:
    raw = raw.replace("\ufeff", "")
    cleaned: list[str] = []
    for line in raw.splitlines():
        s = line.strip().lower()
        if s.startswith("attribute vb_") or s.startswith("version ") or s in {"begin", "end"}:
            continue
        cleaned.append(line)
    return "\r\n".join(cleaned).rstrip() + "\r\n"


def _install_modules(workbook: Path) -> None:
    if win32 is None:
        raise RuntimeError("pywin32 is required for VBA repairs")

    excel = win32.DispatchEx("Excel.Application")  # type: ignore[attr-defined]
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    wb = excel.Workbooks.Open(str(workbook), ReadOnly=False, UpdateLinks=False, AddToMru=False)
    try:
        vbproj = wb.VBProject  # type: ignore[attr-defined]
        comps = vbproj.VBComponents

        for name in ("AutoTraderAdvanced", "cDashboardWatcher"):
            for idx in range(comps.Count, 0, -1):
                comp = comps.Item(idx)
                if comp.Name.lower() == name.lower():
                    comps.Remove(comp)
                    break

        std = comps.Add(1)
        std.Name = "AutoTraderAdvanced"
        std.CodeModule.AddFromString(_sanitize_module_text(AUTO_SRC.read_text(encoding="utf-8", errors="ignore")))

        cls = comps.Add(2)
        cls.Name = "cDashboardWatcher"
        cls.CodeModule.AddFromString(
            "\r\n".join(
                [
                    "Option Explicit",
                    "Public WithEvents App As Application",
                    "",
                    "Private Sub Class_Initialize()",
                    "    Set App = Application",
                    "End Sub",
                    "",
                    "Private Sub Class_Terminate()",
                    "    Set App = Nothing",
                    "End Sub",
                    "",
                    "Private Sub App_SheetCalculate(ByVal Sh As Object)",
                    "    On Error Resume Next",
                    "    Application.Run \"AutoTraderAdvanced.OnDashboardCalculate\", Sh",
                    "End Sub",
                    "",
                ]
            )
        )

        wb.Save()
    finally:
        wb.Close(SaveChanges=True)
        excel.Quit()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel", default=str(DEFAULT_WORKBOOK))
    args = parser.parse_args()

    workbook = Path(args.excel).resolve()
    if not workbook.exists():
        raise SystemExit(f"Workbook not found: {workbook}")

    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = workbook.with_name(f"{workbook.stem}_backup_{ts}{workbook.suffix}")
    shutil.copy2(workbook, backup)

    run(["scripts/repair_asagake_dashboard.py", "--excel", str(workbook)])
    _install_modules(workbook)
    run(["scripts/restore_dashboard_formulas.py", "--excel", str(workbook)])

    result = {
        "backup": str(backup),
        "excel": str(workbook),
        "modules_rebuilt": True,
        "timestamp": ts,
    }
    print(json.dumps(result))


if __name__ == "__main__":
    main()
