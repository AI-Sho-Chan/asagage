from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

FORBIDDEN_IMPORT = re.compile(r"from\s+openpyxl\s+import|import\s+openpyxl", re.I)
PAT_SAVE = re.compile(r"\.save\s*\(\s*[^\)]*\)\s*$", re.I | re.M)
PAT_ASSIGN = re.compile(r"\[[\"'][A-Z]+\d+[\"']\]\s*=", re.I)
PAT_XLSM = re.compile(r"SHINSOKU\.xlsm|\.xlsm", re.I)

def file_has_forbidden(text: str) -> bool:
    # Allow openpyxl read-only exports. Flag only if:
    #  - openpyxl is imported, AND
    #  - target is .xlsm, AND
    #  - there is evidence of write (wb.save or cell assignment), OR load_workbook without read_only=True
    if not FORBIDDEN_IMPORT.search(text):
        return False
    if not PAT_XLSM.search(text):
        return False
    if PAT_SAVE.search(text) or PAT_ASSIGN.search(text):
        return True
    return False

def main() -> int:
    offenders: list[str] = []
    for p in ROOT.rglob("*.py"):
        # Skip virtual environments and git internals
        parts = {q.lower() for q in p.parts}
        if any(x in parts for x in {".venv", ".git", "venv", "env"}):
            continue
        try:
            text = p.read_text(errors="ignore")
        except Exception:
            continue
        if file_has_forbidden(text):
            offenders.append(str(p.relative_to(ROOT)))
    if offenders:
        print("FORBIDDEN: openpyxl-based .xlsm write detected in:")
        for o in offenders:
            print(" -", o)
        print("Use VBA/COM scripts instead (see AGENTS.md and docs/codex.md).")
        return 1
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
