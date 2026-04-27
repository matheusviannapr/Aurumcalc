from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd


LOAD_PROFILE_ROOT = Path("data/load_profiles")


@dataclass
class BillingGroupAInput:
    concessionaria: str
    modalidade_tarifaria: str  # azul|verde
    subgrupo: str
    mes_referencia: str  # YYYY-MM
    dias_ciclo: int
    ponta_inicio_hora: int
    ponta_fim_hora: int
    feriados_nacionais: List[str]
    energia_ponta_kwh: Optional[float] = None
    energia_fora_ponta_kwh: Optional[float] = None
    demanda_medida_ponta_kw: Optional[float] = None
    demanda_medida_fora_ponta_kw: Optional[float] = None
    demanda_contratada_ponta_kw: Optional[float] = None
    demanda_contratada_fora_ponta_kw: Optional[float] = None
    tarifa_energia_ponta: float = 0.0
    tarifa_energia_fora_ponta: float = 0.0
    tarifa_demanda_ponta: float = 0.0
    tarifa_demanda_fora_ponta: float = 0.0
    impostos_ajustes: float = 0.0


@dataclass
class BillingGroupBInput:
    concessionaria: str
    modalidade: str  # convencional|branca
    mes_referencia: str  # YYYY-MM
    dias_ciclo: int
    energia_total_kwh: float
    energia_ponta_kwh: Optional[float] = None
    energia_intermediaria_kwh: Optional[float] = None
    energia_fora_ponta_kwh: Optional[float] = None
    tarifa_energia: float = 0.0
    tarifa_energia_ponta: float = 0.0
    tarifa_energia_intermediaria: float = 0.0
    tarifa_energia_fora_ponta: float = 0.0
    area_disponivel_m2: Optional[float] = None
    perfil_consumo: Optional[str] = None


@dataclass
class ProfileCandidate:
    metadata: Dict[str, Any]
    normalized_vector: np.ndarray


@dataclass
class CalibrationResult:
    perfil_escolhido_id: str
    fator_aplicado: float
    curva_calibrada_kw: np.ndarray
    erros_percentuais: Dict[str, Optional[float]]
    score: float
    alertas: List[str]


def _safe_div(num: float, den: float) -> Optional[float]:
    if den is None or den == 0:
        return None
    return num / den


def list_profiles_by_sector(setor: str, root: Path = LOAD_PROFILE_ROOT) -> List[Dict[str, Any]]:
    index_path = root / "profiles_index.json"
    if not index_path.exists():
        return []
    data = json.loads(index_path.read_text(encoding="utf-8"))
    return [p for p in data.get("profiles", []) if p.get("setor") == setor]


def _ensure_8760(values: Sequence[float], resolution: str) -> np.ndarray:
    arr = np.asarray(values, dtype=float)
    if len(arr) == 8760:
        return arr
    if len(arr) == 24 or resolution.lower().startswith("24"):
        return np.tile(arr, 365)
    raise ValueError(f"Perfil inválido: tamanho={len(arr)} resolução={resolution}")


def load_profile(profile_path: str | Path, root: Path = LOAD_PROFILE_ROOT) -> ProfileCandidate:
    fp = (root / profile_path) if not Path(profile_path).is_absolute() else Path(profile_path)
    data = json.loads(fp.read_text(encoding="utf-8"))
    vector = data.get("hourly_vector_normalized") or data.get("hourly_vector_kw")
    if vector is None:
        raise ValueError(f"Perfil sem vetor horário: {fp}")
    curve = _ensure_8760(vector, data.get("resolucao_final", "8760h"))
    if np.all(curve == 0):
        raise ValueError("Curva de carga não pode ser toda zero.")
    if data.get("hourly_vector_kw") is None:
        curve = curve / np.sum(curve)
    return ProfileCandidate(metadata=data, normalized_vector=curve)


def expand_month_mask(index: pd.DatetimeIndex, dias_funcionamento_semana: int = 7,
                      horario_funcionamento: Tuple[int, int] = (0, 24)) -> np.ndarray:
    start_h, end_h = horario_funcionamento
    weekdays_enabled = set(range(min(max(dias_funcionamento_semana, 0), 7)))
    mask = []
    for ts in index:
        ok_day = ts.weekday() in weekdays_enabled
        ok_hour = start_h <= ts.hour < end_h
        mask.append(1.0 if (ok_day and ok_hour) else 0.0)
    return np.asarray(mask, dtype=float)


def classify_tariff_periods(
    index: pd.DatetimeIndex,
    ponta_inicio_hora: int,
    ponta_fim_hora: int,
    feriados_nacionais: Optional[Iterable[str]] = None,
    modalidade_grupo_b: Optional[str] = None,
) -> np.ndarray:
    holidays = {pd.Timestamp(h).date() for h in (feriados_nacionais or [])}
    periods: List[str] = []
    for ts in index:
        is_weekend_or_holiday = ts.weekday() >= 5 or ts.date() in holidays
        if is_weekend_or_holiday:
            periods.append("fora_ponta")
            continue
        hour = ts.hour
        if modalidade_grupo_b == "branca":
            if ponta_inicio_hora <= hour < ponta_fim_hora:
                periods.append("ponta")
            elif (ponta_inicio_hora - 1) <= hour < ponta_inicio_hora or ponta_fim_hora <= hour < (ponta_fim_hora + 1):
                periods.append("intermediario")
            else:
                periods.append("fora_ponta")
        else:
            periods.append("ponta" if ponta_inicio_hora <= hour < ponta_fim_hora else "fora_ponta")
    return np.asarray(periods)


def month_hour_index(month_str: str, days_in_cycle: int) -> pd.DatetimeIndex:
    year, month = map(int, month_str.split("-"))
    return pd.date_range(datetime(year, month, 1), periods=days_in_cycle * 24, freq="H")


def calibrate_load_profile(
    candidate_profiles: List[ProfileCandidate],
    billing_data: BillingGroupAInput | BillingGroupBInput,
    customer_type: str,
    dias_funcionamento_semana: int = 7,
    horario_funcionamento: Tuple[int, int] = (0, 24),
    tolerancia_erro_percentual: float = 5.0,
) -> CalibrationResult:
    idx = month_hour_index(billing_data.mes_referencia, billing_data.dias_ciclo)
    periods = classify_tariff_periods(
        idx,
        getattr(billing_data, "ponta_inicio_hora", 18),
        getattr(billing_data, "ponta_fim_hora", 21),
        getattr(billing_data, "feriados_nacionais", []),
        getattr(billing_data, "modalidade", None),
    )
    operation_mask = expand_month_mask(idx, dias_funcionamento_semana, horario_funcionamento)

    best: Optional[CalibrationResult] = None
    for candidate in candidate_profiles:
        full_year = candidate.normalized_vector
        month_curve = full_year[: len(idx)] * operation_mask

        ponta_mask = periods == "ponta"
        fora_mask = periods != "ponta"

        e_ponta_std = float(np.sum(month_curve[ponta_mask]))
        e_fora_std = float(np.sum(month_curve[fora_mask]))
        e_total_std = float(np.sum(month_curve))
        d_ponta_std = float(np.max(month_curve[ponta_mask])) if np.any(ponta_mask) else 0.0
        d_fora_std = float(np.max(month_curve[fora_mask])) if np.any(fora_mask) else 0.0

        if isinstance(billing_data, BillingGroupAInput):
            target_ponta = billing_data.energia_ponta_kwh
            target_total = (billing_data.energia_ponta_kwh or 0) + (billing_data.energia_fora_ponta_kwh or 0)
            target_fora = billing_data.energia_fora_ponta_kwh
            target_dp = billing_data.demanda_medida_ponta_kw
            target_df = billing_data.demanda_medida_fora_ponta_kw
        else:
            target_ponta = billing_data.energia_ponta_kwh
            target_total = billing_data.energia_total_kwh
            target_fora = billing_data.energia_fora_ponta_kwh
            target_dp = None
            target_df = None

        base = e_ponta_std if (target_ponta and e_ponta_std > 0) else e_total_std
        target_for_factor = target_ponta if (target_ponta and e_ponta_std > 0) else target_total
        factor = float(target_for_factor / base) if base > 0 else 0.0

        calibrated = month_curve * factor
        e_ponta = float(np.sum(calibrated[ponta_mask]))
        e_fora = float(np.sum(calibrated[fora_mask]))
        e_total = float(np.sum(calibrated))
        d_ponta = float(np.max(calibrated[ponta_mask])) if np.any(ponta_mask) else 0.0
        d_fora = float(np.max(calibrated[fora_mask])) if np.any(fora_mask) else 0.0

        errors = {
            "energia_ponta": (abs(e_ponta - target_ponta) / target_ponta * 100.0) if target_ponta else None,
            "energia_fora_ponta": (abs(e_fora - target_fora) / target_fora * 100.0) if target_fora else None,
            "energia_total": (abs(e_total - target_total) / target_total * 100.0) if target_total else None,
            "demanda_ponta": (abs(d_ponta - target_dp) / target_dp * 100.0) if target_dp else None,
            "demanda_fora_ponta": (abs(d_fora - target_df) / target_df * 100.0) if target_df else None,
        }

        weighted = {
            "energia_ponta": 0.35,
            "energia_fora_ponta": 0.25,
            "energia_total": 0.2,
            "demanda_ponta": 0.1,
            "demanda_fora_ponta": 0.1,
        }
        score = 0.0
        for k, w in weighted.items():
            if errors[k] is not None:
                score += errors[k] * w

        alerts = [
            f"Erro alto em {k}: {v:.1f}%"
            for k, v in errors.items()
            if v is not None and v > tolerancia_erro_percentual
        ]

        result = CalibrationResult(
            perfil_escolhido_id=candidate.metadata.get("id", "unknown"),
            fator_aplicado=factor,
            curva_calibrada_kw=calibrated,
            erros_percentuais=errors,
            score=score,
            alertas=alerts,
        )
        if best is None or result.score < best.score:
            best = result

    if best is None:
        raise ValueError(f"Nenhum perfil disponível para calibrar {customer_type}")
    return best


def calculate_hourly_energy_balance(load_kw: Sequence[float], pv_kw: Sequence[float]) -> pd.DataFrame:
    load = np.asarray(load_kw, dtype=float)
    pv = np.asarray(pv_kw, dtype=float)
    n = min(len(load), len(pv))
    load = load[:n]
    pv = pv[:n]
    self_cons = np.minimum(load, pv)
    grid = np.maximum(load - pv, 0)
    exp = np.maximum(pv - load, 0)

    return pd.DataFrame(
        {
            "load_kw": load,
            "pv_kw": pv,
            "self_consumption_kw": self_cons,
            "grid_import_kw": grid,
            "export_kw": exp,
            "load_kwh": load,
            "pv_kwh": pv,
            "self_consumption_kwh": self_cons,
            "grid_import_kwh": grid,
            "export_kwh": exp,
        }
    )


def split_by_tariff_period(balance_df: pd.DataFrame, periods: Sequence[str]) -> Dict[str, float]:
    p = np.asarray(periods[: len(balance_df)])
    ponta = p == "ponta"
    fora = p == "fora_ponta"
    inter = p == "intermediario"

    return {
        "energia_consumida_ponta_antes_fv": float(balance_df.loc[ponta, "load_kwh"].sum()),
        "energia_consumida_fora_ponta_antes_fv": float(balance_df.loc[fora, "load_kwh"].sum()),
        "energia_importada_ponta_apos_fv": float(balance_df.loc[ponta, "grid_import_kwh"].sum()),
        "energia_importada_fora_ponta_apos_fv": float(balance_df.loc[fora, "grid_import_kwh"].sum()),
        "energia_autoconsumida_ponta": float(balance_df.loc[ponta, "self_consumption_kwh"].sum()),
        "energia_autoconsumida_fora_ponta": float(balance_df.loc[fora, "self_consumption_kwh"].sum()),
        "excedente_injetado_ponta": float(balance_df.loc[ponta, "export_kwh"].sum()),
        "excedente_injetado_fora_ponta": float(balance_df.loc[fora, "export_kwh"].sum()),
        "energia_importada_intermediario_apos_fv": float(balance_df.loc[inter, "grid_import_kwh"].sum()),
        "demanda_max_ponta_antes": float(balance_df.loc[ponta, "load_kw"].max()) if np.any(ponta) else 0.0,
        "demanda_max_ponta_depois": float(balance_df.loc[ponta, "grid_import_kw"].max()) if np.any(ponta) else 0.0,
        "demanda_max_fora_ponta_antes": float(balance_df.loc[fora, "load_kw"].max()) if np.any(fora) else 0.0,
        "demanda_max_fora_ponta_depois": float(balance_df.loc[fora, "grid_import_kw"].max()) if np.any(fora) else 0.0,
    }


def compute_economic_summary(split: Dict[str, float], billing_data: BillingGroupAInput | BillingGroupBInput,
                            compensacao_excedente_rs_kwh: float = 0.0,
                            economia_load_shift: float = 0.0,
                            economia_peak_shaving: float = 0.0) -> Dict[str, float]:
    econ_ponta = (split["energia_consumida_ponta_antes_fv"] - split["energia_importada_ponta_apos_fv"]) * getattr(
        billing_data, "tarifa_energia_ponta", 0.0
    )
    econ_fora = (split["energia_consumida_fora_ponta_antes_fv"] - split["energia_importada_fora_ponta_apos_fv"]) * getattr(
        billing_data, "tarifa_energia_fora_ponta", getattr(billing_data, "tarifa_energia", 0.0)
    )
    red_dem_p = (split["demanda_max_ponta_antes"] - split["demanda_max_ponta_depois"]) * getattr(billing_data, "tarifa_demanda_ponta", 0.0)
    red_dem_f = (split["demanda_max_fora_ponta_antes"] - split["demanda_max_fora_ponta_depois"]) * getattr(billing_data, "tarifa_demanda_fora_ponta", 0.0)

    self_total = split["energia_autoconsumida_ponta"] + split["energia_autoconsumida_fora_ponta"]
    load_total = split["energia_consumida_ponta_antes_fv"] + split["energia_consumida_fora_ponta_antes_fv"]
    pv_used = self_total + split["excedente_injetado_ponta"] + split["excedente_injetado_fora_ponta"]

    total = econ_ponta + econ_fora + red_dem_p + red_dem_f + economia_load_shift + economia_peak_shaving
    excedente = split["excedente_injetado_ponta"] + split["excedente_injetado_fora_ponta"]
    compensacao = excedente * compensacao_excedente_rs_kwh
    total += compensacao

    return {
        "economia_energia_ponta": econ_ponta,
        "economia_energia_fora_ponta": econ_fora,
        "economia_demanda_ponta": red_dem_p,
        "economia_demanda_fora_ponta": red_dem_f,
        "economia_autoconsumo_instantaneo": econ_ponta + econ_fora,
        "economia_compensacao_excedente": compensacao,
        "economia_load_shifting": economia_load_shift,
        "economia_peak_shaving": economia_peak_shaving,
        "economia_total_mensal": total,
        "economia_anual_projetada": total * 12,
        "percentual_autoconsumo": (self_total / pv_used * 100.0) if pv_used > 0 else 0.0,
        "percentual_simultaneidade": (self_total / load_total * 100.0) if load_total > 0 else 0.0,
        "reducao_demanda_maxima": split["demanda_max_ponta_antes"] - split["demanda_max_ponta_depois"],
        "energia_excedente": excedente,
    }


def simulate_load_shifting(
    load_kw: Sequence[float],
    periods: Sequence[str],
    percentual_flexivel: float,
    limite_energia_deslocavel_dia_kwh: float,
    limite_potencia_deslocavel_kw: float,
    eficiencia: float = 1.0,
) -> Dict[str, Any]:
    load = np.asarray(load_kw, dtype=float).copy()
    p = np.asarray(periods[: len(load)])
    shifted = load.copy()

    daily_hours = len(load) // 30 if len(load) >= 720 else 24
    for d_start in range(0, len(load), daily_hours):
        sl = slice(d_start, min(d_start + daily_hours, len(load)))
        day_p = p[sl]
        day_load = shifted[sl].copy()

        src_idx = np.where(day_p == "ponta")[0]
        dst_idx = np.where(day_p == "fora_ponta")[0]
        if len(src_idx) == 0 or len(dst_idx) == 0:
            continue

        potential = np.minimum(day_load[src_idx] * percentual_flexivel, limite_potencia_deslocavel_kw)
        energy_to_move = min(float(np.sum(potential)), limite_energia_deslocavel_dia_kwh)
        if energy_to_move <= 0:
            continue

        for i in src_idx:
            reduc = min(shifted[sl][i] * percentual_flexivel, limite_potencia_deslocavel_kw, energy_to_move)
            shifted[sl][i] -= reduc
            energy_to_move -= reduc
            if energy_to_move <= 0:
                break

        energy_add = (min(float(np.sum(potential)), limite_energia_deslocavel_dia_kwh) * eficiencia)
        per_hour = energy_add / len(dst_idx)
        shifted[sl][dst_idx] += per_hour

    return {"shifted_load_kw": shifted, "energia_total_antes": float(np.sum(load)), "energia_total_depois": float(np.sum(shifted))}


def simulate_peak_shaving(
    net_load_kw: Sequence[float],
    demand_target_kw: float,
    battery_power_kw: float,
    battery_capacity_kwh: float,
    soc_min: float = 0.1,
    soc_max: float = 0.9,
    roundtrip_efficiency: float = 0.9,
) -> Dict[str, Any]:
    load = np.asarray(net_load_kw, dtype=float)
    clipped = load.copy()
    soc = np.zeros_like(load)
    soc_energy = battery_capacity_kwh * soc_max
    min_energy = battery_capacity_kwh * soc_min
    max_energy = battery_capacity_kwh * soc_max

    batt_power = np.zeros_like(load)
    for i, l in enumerate(load):
        if l > demand_target_kw:
            needed = l - demand_target_kw
            discharge = min(needed, battery_power_kw, soc_energy - min_energy)
            clipped[i] = l - discharge
            soc_energy -= discharge
            batt_power[i] = discharge
        else:
            room = max_energy - soc_energy
            charge = min(demand_target_kw - l, battery_power_kw, room)
            soc_energy += charge * roundtrip_efficiency
            batt_power[i] = -charge
        soc[i] = soc_energy / battery_capacity_kwh if battery_capacity_kwh > 0 else soc_min

    return {
        "net_load_after_peak_shaving_kw": clipped,
        "battery_power_kw": batt_power,
        "soc": soc,
        "demanda_antes": float(np.max(load)),
        "demanda_depois": float(np.max(clipped)),
        "alertas": [] if np.max(clipped) <= demand_target_kw + 1e-6 else ["Meta de demanda não atingida em todas as horas."],
    }


def calculate_area_limited_pv(
    area_disponivel_m2: float,
    modulo_area_m2: float,
    fator_ocupacao: float,
    potencia_modulo_w: float,
    potencia_sistema_kwp: Optional[float] = None,
) -> Dict[str, Any]:
    area_util = area_disponivel_m2 * fator_ocupacao
    max_modulos = int(area_util // modulo_area_m2) if modulo_area_m2 > 0 else 0
    potencia_max_kwp = max_modulos * potencia_modulo_w / 1000.0
    area_usada = max_modulos * modulo_area_m2
    alertas = []
    if potencia_sistema_kwp and potencia_sistema_kwp > potencia_max_kwp:
        alertas.append("Sistema excede limite de área disponível.")

    return {
        "numero_max_modulos": max_modulos,
        "potencia_max_dc_kwp": potencia_max_kwp,
        "area_usada_m2": area_usada,
        "percentual_ocupacao_real": (area_usada / area_disponivel_m2 * 100.0) if area_disponivel_m2 > 0 else 0.0,
        "alertas": alertas,
    }
