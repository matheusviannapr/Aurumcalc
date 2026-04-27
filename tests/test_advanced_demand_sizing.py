import math

from src.load_profiles import load_typical_profiles, validate_all_profiles
from src.load_profile_fitting import fit_load_profile_to_bill
from src.tariff_periods import classify_peak_hours
from src.map_area import calculate_polygon_area_m2
from src.self_consumption import analyze_self_consumption
from src.peak_shaving import analyze_peak_shaving
from pv_calculator import aplicar_restricao_area_resultados
import pandas as pd


def test_profiles_have_24_points_and_sum_one():
    profiles = load_typical_profiles()
    assert len(profiles) >= 15
    assert validate_all_profiles() == []


def test_fit_group_b_matches_monthly_energy():
    base = [1 / 24] * 24
    peak = classify_peak_hours(list(range(24)), "18:00", "21:00")
    out = fit_load_profile_to_bill(base, {"energia_total_kwh_mes": 720}, 30, peak, "B")
    assert math.isclose(out["monthly_energy_kwh"], 720, rel_tol=1e-6)


def test_fit_group_a_matches_peak_and_offpeak_energy():
    base = [1 / 24] * 24
    peak = classify_peak_hours(list(range(24)), "18:00", "21:00")
    out = fit_load_profile_to_bill(
        base,
        {
            "energia_ponta_kwh_mes": 300,
            "energia_fora_ponta_kwh_mes": 2100,
            "demanda_ponta_kw": 5,
            "demanda_fora_ponta_kw": 4,
        },
        30,
        peak,
        "A",
    )
    assert math.isclose(out["energy_peak_kwh_estimated"], 300, rel_tol=1e-6)
    assert math.isclose(out["energy_offpeak_kwh_estimated"], 2100, rel_tol=1e-6)


def test_classify_peak_hours_basic():
    periods = classify_peak_hours(list(range(24)), 18, 21)
    assert periods[18] == "ponta"
    assert periods[20] == "ponta"
    assert periods[21] == "fora_ponta"


def test_polygon_area_m2():
    coords = [(-23.0, -46.0), (-23.0, -45.9999), (-23.0001, -45.9999), (-23.0001, -46.0)]
    area = calculate_polygon_area_m2(coords)
    assert area > 100


def test_area_module_restriction():
    df = pd.DataFrame(
        [
            {"sistema_num_total_paineis": 30, "painel_potencia": 550},
            {"sistema_num_total_paineis": 80, "painel_potencia": 550},
        ]
    )
    filtered, meta = aplicar_restricao_area_resultados(df, available_area_m2=100, packing_factor=0.7)
    assert meta["max_modules_by_area"] > 0
    assert int(filtered.iloc[0]["sistema_num_total_paineis"]) <= meta["max_modules_by_area"] or meta["warnings"]


def test_self_consumption_calculation():
    out = analyze_self_consumption([5, 2, 0], [3, 4, 1])
    assert out["total_self_consumption_kwh"] == 5
    assert out["total_grid_export_kwh"] == 3


def test_peak_shaving_calculation():
    periods = ["fora_ponta"] * 24
    periods[18:21] = ["ponta", "ponta", "ponta"]
    out = analyze_peak_shaving([10] * 24, [0] * 18 + [2, 3, 2] + [0] * 3, periods)
    assert out["peak_shaving_kw"] >= 0
