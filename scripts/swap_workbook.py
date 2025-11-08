import sys
from pathlib import Path
import shutil
import time


def main() -> None:
    if len(sys.argv) < 3:
        print("Usage: python scripts/swap_workbook.py <patched_xlsm> <target_xlsm>")
        sys.exit(2)
    src = Path(sys.argv[1]).resolve()
    dst = Path(sys.argv[2]).resolve()
    if not src.exists():
        print("Patched workbook not found:", src)
        sys.exit(3)
    if not dst.exists():
        print("Target workbook not found:", dst)
        sys.exit(4)
    bak = dst.with_name(dst.stem + f"_backup_{time.strftime('%Y%m%d_%H%M%S')}.xlsm")
    try:
        shutil.copy2(dst, bak)
        print("Backup saved:", bak)
    except Exception as e:
        print("BACKUP_ERROR", e)
        sys.exit(5)
    try:
        # Copy patched over target in place
        tmp = dst.with_suffix('.tmp')
        shutil.copy2(src, tmp)
        tmp.replace(dst)
        print("Replaced:", dst)
    except Exception as e:
        print("REPLACE_ERROR. Is Excel open?", e)
        sys.exit(6)


if __name__ == "__main__":
    main()

