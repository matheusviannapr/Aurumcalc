from __future__ import annotations

from typing import Any


def simulate_simplified_load_shifting(
    energy_peak_kwh: float,
    energy_offpeak_kwh: float,
    percentual_carga_deslocavel_ponta: float = 0.0,
    shift_window: str = "madrugada",
    tarifa_ponta: float | None = None,
    tarifa_fora_ponta: float | None = None,
) -> dict[str, Any]:
    pct = max(0.0, min(percentual_carga_deslocavel_ponta, 1.0))
    shifted = max(0.0, energy_peak_kwh) * pct

    peak_after = max(0.0, energy_peak_kwh - shifted)
    offpeak_after = max(0.0, energy_offpeak_kwh + shifted)

    savings = None
    if tarifa_ponta is not None and tarifa_fora_ponta is not None:
        savings = shifted * (float(tarifa_ponta) - float(tarifa_fora_ponta))

    return {
        "shift_window": shift_window,
        "percentual_carga_deslocavel_ponta": pct,
        "shiftable_energy_kwh": shifted,
        "energy_peak_after_shift": peak_after,
        "energy_offpeak_after_shift": offpeak_after,
        "estimated_savings_from_load_shift": savings,
    }
