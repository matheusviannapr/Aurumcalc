from __future__ import annotations

from typing import Sequence

import numpy as np


def analyze_self_consumption(load_kw: Sequence[float], pv_generation_kw: Sequence[float]) -> dict:
    load = np.asarray(load_kw, dtype=float)
    pv = np.asarray(pv_generation_kw, dtype=float)
    n = min(len(load), len(pv))
    load = load[:n]
    pv = pv[:n]

    self_consumption = np.minimum(load, pv)
    grid_import = np.maximum(load - pv, 0)
    grid_export = np.maximum(pv - load, 0)

    total_load = float(np.sum(load))
    total_pv = float(np.sum(pv))
    total_self = float(np.sum(self_consumption))

    return {
        "load_kw": load.tolist(),
        "pv_generation_kw": pv.tolist(),
        "self_consumption_kwh": self_consumption.tolist(),
        "grid_import_kwh": grid_import.tolist(),
        "grid_export_kwh": grid_export.tolist(),
        "total_load_kwh": total_load,
        "total_pv_generation_kwh": total_pv,
        "total_self_consumption_kwh": total_self,
        "total_grid_import_kwh": float(np.sum(grid_import)),
        "total_grid_export_kwh": float(np.sum(grid_export)),
        "self_consumption_ratio": (total_self / total_pv) if total_pv > 0 else 0.0,
        "self_sufficiency_ratio": (total_self / total_load) if total_load > 0 else 0.0,
    }
