import numpy as np
import pandas as pd

from hourly_analysis import (
    BillingGroupAInput,
    calculate_area_limited_pv,
    calculate_hourly_energy_balance,
    calibrate_load_profile,
    classify_tariff_periods,
    compute_economic_summary,
    simulate_load_shifting,
    simulate_peak_shaving,
    split_by_tariff_period,
    ProfileCandidate,
)


def test_classify_tariff_weekend_and_holiday_as_offpeak():
    idx = pd.date_range("2026-01-01", periods=72, freq="h")
    periods = classify_tariff_periods(idx, 18, 21, feriados_nacionais=["2026-01-01"])
    # 01/01 é feriado; sábado/domingo sempre fora ponta
    assert periods[18] == "fora_ponta"
    sat = np.where(idx.weekday == 5)[0][0]
    assert periods[sat + 19] == "fora_ponta"


def test_calibration_single_factor_matches_energy_ponta_close():
    base = np.tile(np.array([1.0] * 24) / 24.0, 365)
    candidate = ProfileCandidate(metadata={"id": "test"}, normalized_vector=base)
    billing = BillingGroupAInput(
        concessionaria="X", modalidade_tarifaria="verde", subgrupo="A4", mes_referencia="2026-01", dias_ciclo=30,
        ponta_inicio_hora=18, ponta_fim_hora=21, feriados_nacionais=[], energia_ponta_kwh=300, energia_fora_ponta_kwh=2100,
    )
    result = calibrate_load_profile([candidate], billing, customer_type="escola")
    assert result.fator_aplicado > 0
    assert result.erros_percentuais["energia_ponta"] is not None
    assert result.erros_percentuais["energia_ponta"] < 1.0


def test_hourly_self_consumption_balance():
    load = [5, 2, 0]
    pv = [3, 4, 1]
    df = calculate_hourly_energy_balance(load, pv)
    assert list(df["self_consumption_kw"]) == [3, 2, 0]
    assert list(df["grid_import_kw"]) == [2, 0, 0]
    assert list(df["export_kw"]) == [0, 2, 1]


def test_split_tariff_and_peak_before_after():
    balance = calculate_hourly_energy_balance([10, 8, 4, 2], [0, 2, 3, 1])
    split = split_by_tariff_period(balance, ["ponta", "ponta", "fora_ponta", "fora_ponta"])
    assert split["energia_consumida_ponta_antes_fv"] == 18
    assert split["demanda_max_ponta_antes"] == 10
    assert split["demanda_max_ponta_depois"] == 10


def test_area_constraint():
    area = calculate_area_limited_pv(100, 2, 0.8, 500, potencia_sistema_kwp=30)
    assert area["numero_max_modulos"] == 40
    assert area["potencia_max_dc_kwp"] == 20
    assert area["alertas"]


def test_load_shifting_energy_conservation_with_efficiency():
    load = np.array([10.0] * 24)
    periods = np.array(["fora_ponta"] * 18 + ["ponta"] * 3 + ["fora_ponta"] * 3)
    out = simulate_load_shifting(load, periods, 0.2, 10.0, 5.0, eficiencia=0.9)
    assert out["energia_total_depois"] <= out["energia_total_antes"]


def test_peak_shaving_reduces_max_demand():
    net = np.array([5, 6, 20, 7, 19, 6], dtype=float)
    out = simulate_peak_shaving(net, demand_target_kw=10, battery_power_kw=8, battery_capacity_kwh=20)
    assert out["demanda_depois"] <= out["demanda_antes"]


def test_economic_summary_basic():
    split = {
        "energia_consumida_ponta_antes_fv": 100,
        "energia_consumida_fora_ponta_antes_fv": 200,
        "energia_importada_ponta_apos_fv": 50,
        "energia_importada_fora_ponta_apos_fv": 120,
        "energia_autoconsumida_ponta": 50,
        "energia_autoconsumida_fora_ponta": 80,
        "excedente_injetado_ponta": 5,
        "excedente_injetado_fora_ponta": 10,
        "demanda_max_ponta_antes": 30,
        "demanda_max_ponta_depois": 20,
        "demanda_max_fora_ponta_antes": 25,
        "demanda_max_fora_ponta_depois": 20,
        "energia_importada_intermediario_apos_fv": 0,
    }
    billing = BillingGroupAInput(
        concessionaria="X", modalidade_tarifaria="verde", subgrupo="A4", mes_referencia="2026-01", dias_ciclo=30,
        ponta_inicio_hora=18, ponta_fim_hora=21, feriados_nacionais=[], tarifa_energia_ponta=1.2,
        tarifa_energia_fora_ponta=0.7, tarifa_demanda_ponta=30, tarifa_demanda_fora_ponta=20,
    )
    econ = compute_economic_summary(split, billing)
    assert econ["economia_total_mensal"] > 0
