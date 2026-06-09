"""Testes do módulo de análise financeira (src/financials.py)."""
import pytest
from src.financials import (
    calcular_co2_evitado,
    calcular_fluxo_caixa,
    calcular_payback_descontado,
    calcular_tir_real,
    calcular_vpn,
    gerar_excel_completo,
)
import pandas as pd


def test_fluxo_caixa_25_anos_correto():
    fluxo = calcular_fluxo_caixa(300_000, 60_000, 0.85, anos=25)
    assert len(fluxo) == 25
    assert fluxo[0]["Ano"] == 1
    assert fluxo[-1]["Ano"] == 25
    # Geração deve diminuir com degradação
    assert fluxo[24]["Geração (kWh)"] < fluxo[0]["Geração (kWh)"]
    # Tarifa deve aumentar com escalada
    assert fluxo[24]["Tarifa (R$/kWh)"] > fluxo[0]["Tarifa (R$/kWh)"]


def test_fluxo_caixa_sem_degradacao_sem_escalada():
    fluxo = calcular_fluxo_caixa(100_000, 50_000, 1.0, degradacao_anual=0.0, escalada_tarifa=0.0, anos=5)
    for row in fluxo:
        assert abs(row["Geração (kWh)"] - 50_000) < 1
        assert abs(row["Tarifa (R$/kWh)"] - 1.0) < 1e-6


def test_vpn_rentavel():
    fluxo = calcular_fluxo_caixa(100_000, 80_000, 1.0, anos=25, taxa_desconto=0.10)
    vpn = calcular_vpn(100_000, fluxo)
    assert vpn > 0, "Projeto muito rentável deve ter VPL positivo"


def test_payback_descontado_identificado():
    fluxo = calcular_fluxo_caixa(100_000, 50_000, 1.0, anos=25, taxa_desconto=0.10)
    pb = calcular_payback_descontado(fluxo)
    assert pb is not None
    assert 1 < pb < 25


def test_payback_descontado_nao_atingido():
    fluxo = calcular_fluxo_caixa(10_000_000, 1, 0.01, anos=25)
    pb = calcular_payback_descontado(fluxo)
    assert pb is None


def test_tir_real_positiva():
    tir = calcular_tir_real(100_000, 50_000, 0.85, anos=25)
    assert tir is not None
    assert tir > 0


def test_tir_real_negativa_projeto_ruim():
    tir = calcular_tir_real(10_000_000, 1_000, 0.50, anos=25)
    assert tir is None or tir < 0


def test_co2_evitado_proporcional():
    co2a = calcular_co2_evitado(1_000_000)
    co2b = calcular_co2_evitado(2_000_000)
    assert abs(co2b["co2_evitado_t"] / co2a["co2_evitado_t"] - 2.0) < 0.01


def test_co2_arvores_equivalentes_positivo():
    co2 = calcular_co2_evitado(500_000)
    assert co2["arvores_equivalentes"] > 0
    assert co2["co2_evitado_kg"] > 0


def test_gerar_excel_completo_retorna_bytes():
    fluxo = calcular_fluxo_caixa(200_000, 40_000, 0.85, anos=25)
    co2 = calcular_co2_evitado(sum(r["Geração (kWh)"] for r in fluxo))
    eco = {
        "economia_anual": 34_000,
        "payback": 5.88,
        "payback_descontado": 7.5,
        "vpn": 150_000,
        "tir_real": 0.18,
        "roi": 80,
    }
    hist = pd.DataFrame({"Mês": ["Jan", "Fev"], "Consumo (kWh)": [1000, 1100]})
    prem = {"Cliente": "Teste", "Local": "SP"}
    dim = {"sistema_potencia_total_w": 50_000}
    data = gerar_excel_completo(hist, dim, fluxo, eco, co2, prem)
    assert isinstance(data, bytes)
    assert len(data) > 5000
