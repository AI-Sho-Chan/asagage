from __future__ import annotations

from pathlib import Path
import shutil
import time


def backup_workbook(src: Path, out_dir: Path | None = None) -> Path:
    if out_dir is None:
        out_dir = src.parent / "excel" / "バックアップ"
    out_dir.mkdir(parents=True, exist_ok=True)
    ts = time.strftime("%Y%m%d_%H%M%S")
    dst = out_dir / f"{src.stem}_backup_{ts}{src.suffix}"
    shutil.copy2(src, dst)
    return dst


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python scripts/backup_workbook.py C:/AI/asagake/SHINSOKU.xlsm [out_dir]")
        raise SystemExit(2)
    src = Path(sys.argv[1])
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else None
    dst = backup_workbook(src, out)
    print("Backup written:", dst)

