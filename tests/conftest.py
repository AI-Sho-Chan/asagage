from __future__ import annotations

import sys
from pathlib import Path

ROOT = (Path(__file__).resolve().parents[1]).resolve()
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

SRC = (Path(__file__).resolve().parents[1] / "src").resolve()
if SRC.exists():
    sys.path.insert(0, str(SRC))
