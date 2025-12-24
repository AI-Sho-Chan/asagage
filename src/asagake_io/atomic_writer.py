from __future__ import annotations

import csv
import os
from pathlib import Path
from typing import Iterable, Mapping, Sequence


def atomic_write_csv(
    path: Path,
    *,
    columns: Sequence[str],
    rows: Iterable[Mapping[str, object]],
    encoding: str = "utf-8-sig",
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + f".{os.getpid()}.tmp")

    try:
        with open(tmp, "w", newline="", encoding=encoding) as f:
            w = csv.DictWriter(f, fieldnames=list(columns), extrasaction="ignore")
            w.writeheader()
            for row in rows:
                w.writerow({k: ("" if v is None else v) for k, v in row.items()})

        os.replace(tmp, path)
    finally:
        try:
            if tmp.exists():
                tmp.unlink()
        except OSError:
            pass
