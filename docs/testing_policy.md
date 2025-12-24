# Testing policy

Goal: keep the default test run fast and deterministic, without accidentally collecting vendored tests
from virtual environments or large third-party dependencies.

## Standard command (unit tests)

Run from the repo root:

`pytest -q`

## What runs

- Only tests under `tests/` are collected (see `pytest.ini:testpaths`).
- We explicitly ignore `.venv/`, `scripts/`, `output/`, `analysis/`, and other large folders
  (see `pytest.ini:norecursedirs`).

## Notes

- If you need to run ad-hoc experiments under `scripts/`, keep filenames away from `test_*.py`
  to avoid accidental collection outside `tests/`.
