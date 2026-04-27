from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Sequence

import numpy as np


@dataclass
class AdjustedLoadProfile:
    profile_name: str
    hourly_kw: list[float]
    monthly_energy_kwh: float
    peak_energy_kwh: float
    offpeak_energy_kwh: float
    peak_demand_kw: float
    offpeak_demand_kw: float
    fit_factors: dict[str, float]
    warnings: list[str]


def _to_24_curve(base_profile_24h: Sequence[float]) -> np.ndarray:
    arr = np.asarray(base_profile_24h, dtype=float)
    if len(arr) != 24:
        raise ValueError("base_profile_24h deve conter 24 valores")
    total = float(np.sum(arr))
    if total <= 0:
        raise ValueError("base_profile_24h inválida (soma <= 0)")
    return arr / total


def fit_load_profile_to_bill(
    base_profile_24h: Sequence[float],
    bill_data: dict[str, Any],
    operation_days: int,
    peak_hours: Sequence[str],
    group_type: str,
) -> dict[str, Any]:
    curve = _to_24_curve(base_profile_24h)
    if len(peak_hours) != 24:
        raise ValueError("peak_hours deve conter 24 rótulos")
    peak_mask = np.asarray([p == "ponta" for p in peak_hours], dtype=bool)
    offpeak_mask = ~peak_mask

    warnings: list[str] = []
    group = group_type.upper()
    op_days = max(1, int(operation_days))

    if group == "B":
        monthly_kwh = float(bill_data.get("energia_total_kwh_mes", 0.0))
        daily_kwh = monthly_kwh / op_days
        load_curve_kw = curve * daily_kwh

        peak_energy = float(np.sum(load_curve_kw[peak_mask]) * op_days)
        offpeak_energy = float(np.sum(load_curve_kw[offpeak_mask]) * op_days)
        peak_demand = float(np.max(load_curve_kw[peak_mask])) if np.any(peak_mask) else 0.0
        offpeak_demand = float(np.max(load_curve_kw[offpeak_mask])) if np.any(offpeak_mask) else 0.0

        result = {
            "load_curve_kw": load_curve_kw.tolist(),
            "monthly_energy_kwh": float(np.sum(load_curve_kw) * op_days),
            "energy_peak_kwh_estimated": peak_energy,
            "energy_offpeak_kwh_estimated": offpeak_energy,
            "demand_peak_kw_estimated": peak_demand,
            "demand_offpeak_kw_estimated": offpeak_demand,
            "fit_factor_peak": daily_kwh,
            "fit_factor_offpeak": daily_kwh,
            "fit_error_peak_percent": 0.0,
            "fit_error_offpeak_percent": 0.0,
            "warnings": warnings,
        }
        return result

    if group != "A":
        raise ValueError("group_type deve ser A ou B")

    target_peak = float(bill_data.get("energia_ponta_kwh_mes", 0.0))
    target_offpeak = float(bill_data.get("energia_fora_ponta_kwh_mes", 0.0))

    daily_peak_base = float(np.sum(curve[peak_mask]))
    daily_offpeak_base = float(np.sum(curve[offpeak_mask]))
    monthly_peak_base = daily_peak_base * op_days
    monthly_offpeak_base = daily_offpeak_base * op_days

    fit_factor_peak = (target_peak / monthly_peak_base) if monthly_peak_base > 0 else 0.0
    fit_factor_offpeak = (target_offpeak / monthly_offpeak_base) if monthly_offpeak_base > 0 else 0.0

    load_curve_kw = np.zeros(24, dtype=float)
    load_curve_kw[peak_mask] = curve[peak_mask] * fit_factor_peak * op_days
    load_curve_kw[offpeak_mask] = curve[offpeak_mask] * fit_factor_offpeak * op_days

    load_curve_kw = load_curve_kw / op_days

    est_peak = float(np.sum(load_curve_kw[peak_mask]) * op_days)
    est_off = float(np.sum(load_curve_kw[offpeak_mask]) * op_days)

    demand_peak = float(np.max(load_curve_kw[peak_mask])) if np.any(peak_mask) else 0.0
    demand_off = float(np.max(load_curve_kw[offpeak_mask])) if np.any(offpeak_mask) else 0.0

    measured_peak = bill_data.get("demanda_ponta_kw")
    measured_off = bill_data.get("demanda_fora_ponta_kw")

    if measured_peak and measured_peak > 0:
        err = abs(demand_peak - float(measured_peak)) / float(measured_peak) * 100
        if err > 25:
            warnings.append("Demanda estimada de ponta incompatível com demanda medida.")
    if measured_off and measured_off > 0:
        err = abs(demand_off - float(measured_off)) / float(measured_off) * 100
        if err > 25:
            warnings.append("Demanda estimada fora ponta incompatível com demanda medida.")

    return {
        "load_curve_kw": load_curve_kw.tolist(),
        "monthly_energy_kwh": est_peak + est_off,
        "energy_peak_kwh_estimated": est_peak,
        "energy_offpeak_kwh_estimated": est_off,
        "demand_peak_kw_estimated": demand_peak,
        "demand_offpeak_kw_estimated": demand_off,
        "fit_factor_peak": fit_factor_peak,
        "fit_factor_offpeak": fit_factor_offpeak,
        "fit_error_peak_percent": 0.0 if target_peak <= 0 else abs(est_peak - target_peak) / target_peak * 100,
        "fit_error_offpeak_percent": 0.0 if target_offpeak <= 0 else abs(est_off - target_offpeak) / target_offpeak * 100,
        "warnings": warnings,
    }


def adjusted_profile_from_fit(profile_name: str, fit_result: dict[str, Any]) -> AdjustedLoadProfile:
    return AdjustedLoadProfile(
        profile_name=profile_name,
        hourly_kw=[float(v) for v in fit_result["load_curve_kw"]],
        monthly_energy_kwh=float(fit_result["monthly_energy_kwh"]),
        peak_energy_kwh=float(fit_result["energy_peak_kwh_estimated"]),
        offpeak_energy_kwh=float(fit_result["energy_offpeak_kwh_estimated"]),
        peak_demand_kw=float(fit_result["demand_peak_kw_estimated"]),
        offpeak_demand_kw=float(fit_result["demand_offpeak_kw_estimated"]),
        fit_factors={
            "peak": float(fit_result["fit_factor_peak"]),
            "offpeak": float(fit_result["fit_factor_offpeak"]),
        },
        warnings=list(fit_result.get("warnings", [])),
    )
