from __future__ import annotations

from pathlib import Path


def test_vba_sources_do_not_start_with_bom() -> None:
    """
    Guardrail: VBA .bas sources must not start with a BOM.

    A UTF-8 BOM (EF BB BF) can become an invisible character when imported into VBE,
    causing compile errors and forcing manual cleanup. This test makes such regressions
    fail fast in CI.
    """

    repo_root = Path(__file__).resolve().parents[1]
    bas_files = sorted((repo_root / "excel").glob("*.bas"))
    assert bas_files, "expected at least one .bas file under excel/"

    bad: list[str] = []
    for path in bas_files:
        raw = path.read_bytes()
        if raw.startswith(b"\xEF\xBB\xBF") or raw.startswith(b"\xFF\xFE") or raw.startswith(b"\xFE\xFF"):
            bad.append(str(path.relative_to(repo_root)))

    assert not bad, f"VBA sources must not start with a BOM: {bad}"

