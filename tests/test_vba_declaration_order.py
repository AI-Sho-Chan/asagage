from __future__ import annotations

import re
from pathlib import Path


def test_msv1_consts_appear_before_first_procedure() -> None:
    """
    Regression guard:
    VBA module-level declarations must not appear after any End Sub/Function block.

    Specifically, MSV1_* constants must stay in the module declaration section
    (before the first procedure) to avoid VBE compile errors like:
    「End Sub...以降にはコメントのみが記述できます」.
    """

    repo_root = Path(__file__).resolve().parents[1]
    path = repo_root / "excel" / "AutoTraderAdvanced.bas"
    text = path.read_text(encoding="utf-8")
    lines = text.splitlines()

    proc_re = re.compile(r"^\s*(Public|Private)\s+(Sub|Function)\b", re.IGNORECASE)
    ms_const_re = re.compile(r"^\s*(Public|Private)\s+Const\s+MSV1_", re.IGNORECASE)

    first_proc_idx = None
    for idx, line in enumerate(lines):
        if line.lstrip().startswith("'"):
            continue
        if proc_re.search(line):
            first_proc_idx = idx
            break

    assert first_proc_idx is not None, "expected at least one procedure in AutoTraderAdvanced.bas"

    bad_lines: list[str] = []
    for idx, line in enumerate(lines):
        if idx <= first_proc_idx:
            continue
        if ms_const_re.search(line):
            bad_lines.append(f"{idx+1}:{line}")

    assert not bad_lines, "MSV1_ Const must appear before first procedure:\n" + "\n".join(bad_lines)

