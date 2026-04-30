from __future__ import annotations

import csv
from pathlib import Path


def export_csv(path: Path, headers: list[str], rows: list[list[object]]) -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)
    return str(path)
