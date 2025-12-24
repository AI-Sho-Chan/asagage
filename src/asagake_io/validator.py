from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence

from .csv_schemas import schema_for_version


@dataclass(frozen=True)
class ValidationError:
    message: str
    row: Optional[int] = None
    column: Optional[str] = None


def validate_header(
    *,
    header: Sequence[str],
    schema_version: str,
    allow_extra: bool = True,
) -> List[ValidationError]:
    errors: List[ValidationError] = []
    if not header:
        return [ValidationError("CSV header is empty")]
    if header[0] != "schema_version":
        errors.append(
            ValidationError(
                "First column must be schema_version",
                row=0,
                column=header[0] if header else None,
            )
        )

    schema = schema_for_version(schema_version)
    expected = [c.name for c in schema]
    expected_set = set(expected)
    header_set = set(header)

    missing = [c.name for c in schema if c.required and c.name not in header_set]
    for m in missing:
        errors.append(ValidationError("Missing required column", row=0, column=m))

    if not allow_extra:
        extra = [h for h in header if h not in expected_set]
        for h in extra:
            errors.append(ValidationError("Unexpected column", row=0, column=h))

    return errors


def validate_csv(
    path: Path,
    *,
    schema_version: str,
    allow_extra: bool = True,
    max_rows: int = 2000,
) -> List[ValidationError]:
    if not path.exists():
        return [ValidationError(f"CSV not found: {path.as_posix()}")]

    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f)
        try:
            header = next(r)
        except StopIteration:
            return [ValidationError("CSV is empty")]

        errors = validate_header(header=header, schema_version=schema_version, allow_extra=allow_extra)
        if errors:
            return errors

        schema = schema_for_version(schema_version)
        types: Dict[str, str] = {c.name: c.typ for c in schema}
        required = {c.name for c in schema if c.required}

        for i, row in enumerate(r, start=2):
            if i > max_rows:
                break
            if not row:
                continue
            if len(row) < len(header):
                # trailing empty fields may be omitted; csv.reader keeps len==header unless malformed
                pass
            row_map = {header[j]: (row[j] if j < len(row) else "") for j in range(len(header))}

            if row_map.get("schema_version") and row_map["schema_version"] != schema_version:
                return [ValidationError("schema_version mismatch in row", row=i, column="schema_version")]

            for col in required:
                if str(row_map.get(col, "")).strip() == "":
                    errors.append(ValidationError("Required value is empty", row=i, column=col))

            for col, typ in types.items():
                if col not in row_map:
                    continue
                val = str(row_map.get(col, "")).strip()
                if val == "":
                    continue
                if typ == "int":
                    try:
                        int(float(val))
                    except ValueError:
                        errors.append(ValidationError("Invalid int", row=i, column=col))
                elif typ == "float":
                    try:
                        float(val)
                    except ValueError:
                        errors.append(ValidationError("Invalid float", row=i, column=col))
                elif typ == "bool":
                    if val not in {"0", "1", "true", "false", "True", "False"}:
                        errors.append(ValidationError("Invalid bool", row=i, column=col))
                else:
                    # str: no check
                    pass

    return errors
