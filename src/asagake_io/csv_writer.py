from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Mapping, Optional, Sequence


def _file_has_header(path: Path) -> bool:
    if not path.exists():
        return False
    try:
        return path.stat().st_size > 0
    except OSError:
        return False


@dataclass
class AppendOnlyCsvWriter:
    path: Path
    columns: Sequence[str]
    schema_version: str
    encoding_new_file: str = "utf-8-sig"
    encoding_append: str = "utf-8"

    def append_rows(self, rows: Iterable[Mapping[str, object]]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)

        is_new = not _file_has_header(self.path)
        encoding = self.encoding_new_file if is_new else self.encoding_append
        mode = "w" if is_new else "a"

        with open(self.path, mode, newline="", encoding=encoding) as f:
            w = csv.DictWriter(f, fieldnames=list(self.columns), extrasaction="ignore")
            if is_new:
                w.writeheader()
            for row in rows:
                row2 = dict(row)
                row2.setdefault("schema_version", self.schema_version)
                w.writerow({k: ("" if v is None else v) for k, v in row2.items()})


@dataclass
class DecisionTraceWriter:
    writer: AppendOnlyCsvWriter
    run_id: str
    env: str
    engine: str
    engine_version: str
    trade_date: str
    source: str
    _event_seq: int = 0

    def next_event_seq(self) -> int:
        self._event_seq += 1
        return self._event_seq

    def append_event(self, event: Mapping[str, object]) -> None:
        event_row = dict(event)
        event_row.setdefault("schema_version", self.writer.schema_version)
        event_row.setdefault("run_id", self.run_id)
        event_row.setdefault("env", self.env)
        event_row.setdefault("engine", self.engine)
        event_row.setdefault("engine_version", self.engine_version)
        event_row.setdefault("trade_date", self.trade_date)
        event_row.setdefault("source", self.source)
        if not event_row.get("event_seq"):
            event_row["event_seq"] = self.next_event_seq()
        self.writer.append_rows([event_row])


def make_append_only_writer(
    path: Path,
    *,
    schema_version: str,
    columns: Sequence[str],
) -> AppendOnlyCsvWriter:
    return AppendOnlyCsvWriter(path=path, columns=columns, schema_version=schema_version)
