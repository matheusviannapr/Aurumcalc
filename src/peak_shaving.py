from __future__ import annotations

from typing import Sequence

import numpy as np


def analyze_peak_shaving(load_kw: Sequence[float], pv_generation_kw: Sequence[float], periods_24h: Sequence[str]) -> dict:
    load = np.asarray(load_kw, dtype=float)
    pv = np.asarray(pv_generation_kw, dtype=float)
    n = min(len(load), len(pv), len(periods_24h))
    load = load[:n]
    pv = pv[:n]
    periods = np.asarray(periods_24h[:n])

    net = np.maximum(load - pv, 0)
    peak_mask = periods == "ponta"
    offpeak_mask = ~peak_mask

    original_peak = float(np.max(load)) if n else 0.0
    net_peak = float(np.max(net)) if n else 0.0

    return {
        "net_load_kw": net.tolist(),
        "original_peak_demand_kw": original_peak,
        "net_peak_demand_kw": net_peak,
        "peak_shaving_kw": original_peak - net_peak,
        "peak_shaving_percent": ((original_peak - net_peak) / original_peak * 100.0) if original_peak > 0 else 0.0,
        "original_peak_demand_peak_hours_kw": float(np.max(load[peak_mask])) if np.any(peak_mask) else 0.0,
        "net_peak_demand_peak_hours_kw": float(np.max(net[peak_mask])) if np.any(peak_mask) else 0.0,
        "peak_shaving_peak_hours_kw": (
            float(np.max(load[peak_mask])) - float(np.max(net[peak_mask]))
        )
        if np.any(peak_mask)
        else 0.0,
        "original_peak_demand_offpeak_hours_kw": float(np.max(load[offpeak_mask])) if np.any(offpeak_mask) else 0.0,
        "net_peak_demand_offpeak_hours_kw": float(np.max(net[offpeak_mask])) if np.any(offpeak_mask) else 0.0,
        "peak_shaving_offpeak_hours_kw": (
            float(np.max(load[offpeak_mask])) - float(np.max(net[offpeak_mask]))
        )
        if np.any(offpeak_mask)
        else 0.0,
    }
