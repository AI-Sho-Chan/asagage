#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from repair_asagake_dashboard import build_dashboard  # type: ignore


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel", default="C:/AI/asagake/ASAGAKE.xlsm")
    args = parser.parse_args()

    excel_path = Path(args.excel).resolve()
    if not excel_path.exists():
        raise SystemExit(f"Workbook not found: {excel_path}")

    build_dashboard(excel_path)
    print("Dashboard rebuilt via repair_asagake_dashboard (formulas refreshed).")


if __name__ == "__main__":
    main()
