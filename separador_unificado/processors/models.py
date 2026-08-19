from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class ProcessResult:
    output_path: Path
    total_rows: int = 0
    entradas: int = 0
    saidas: int = 0
    banco: str = ""

