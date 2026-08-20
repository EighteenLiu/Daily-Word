from __future__ import annotations

import tempfile
from pathlib import Path


WORKSPACE = Path(__file__).resolve().parents[1]
TEMP_ROOT = WORKSPACE / "tmp"


def temporary_directory(prefix: str) -> tempfile.TemporaryDirectory:
    TEMP_ROOT.mkdir(parents=True, exist_ok=True)
    return tempfile.TemporaryDirectory(prefix=prefix, dir=TEMP_ROOT)
