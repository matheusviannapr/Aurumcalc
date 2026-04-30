from __future__ import annotations

from pathlib import Path


def ensure_assets_dir(output_dir: str) -> Path:
    p = Path(output_dir) / "assets"
    p.mkdir(parents=True, exist_ok=True)
    return p
