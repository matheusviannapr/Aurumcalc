"""
Módulo de análise financeira para sistemas fotovoltaicos.
Cobre: fluxo de caixa 25 anos, VPL, TIR real, payback descontado, CO₂ evitado, Excel export.
"""
from __future__ import annotations

import io
from typing import Optional

import pandas as pd

# Fator de emissão do SIN — MCTIC 2023 (tCO₂eq/MWh)
FATOR_EMISSAO_SIN_TCO2_MWH = 0.0839

# Intensidade carbônica por modo de transporte para equivalência ilustrativa
ARVORES_POR_TCO2_ANO = 10  # aproximação: 1 árvore absorve ~100 kg CO₂/ano


# ---------------------------------------------------------------------------
# Fluxo de caixa e indicadores
# ---------------------------------------------------------------------------

def calcular_fluxo_caixa(
    capex: float,
    geracao_ano1_kwh: float,
    tarifa_base: float,
    degradacao_anual: float = 0.006,
    escalada_tarifa: float = 0.05,
    anos: int = 25,
    taxa_desconto: float = 0.10,
    opex_anual_rs: float = 0.0,
) -> list[dict]:
    """
    Retorna lista de dicts com o fluxo de caixa ano a ano.
    Considera degradação linear dos painéis e escalada tarifária composta.
    """
    rows = []
    vpn_acumulado = -capex
    for ano in range(1, anos + 1):
        geracao = geracao_ano1_kwh * ((1 - degradacao_anual) ** (ano - 1))
        tarifa = tarifa_base * ((1 + escalada_tarifa) ** (ano - 1))
        economia_bruta = geracao * tarifa
        fluxo_liquido = economia_bruta - opex_anual_rs
        fator_desconto = (1 + taxa_desconto) ** ano
        fluxo_descontado = fluxo_liquido / fator_desconto
        vpn_acumulado += fluxo_descontado
        rows.append({
            "Ano": ano,
            "Geração (kWh)": round(geracao, 0),
            "Tarifa (R$/kWh)": round(tarifa, 4),
            "Economia bruta (R$)": round(economia_bruta, 2),
            "Fluxo líquido (R$)": round(fluxo_liquido, 2),
            "Fluxo descontado (R$)": round(fluxo_descontado, 2),
            "VPL acumulado (R$)": round(vpn_acumulado, 2),
        })
    return rows


def calcular_payback_descontado(fluxo: list[dict]) -> Optional[float]:
    """Retorna o ano (interpolado) em que o VPL acumulado cruza zero."""
    prev_vpn = None
    for row in fluxo:
        vpn = row["VPL acumulado (R$)"]
        ano = row["Ano"]
        if vpn >= 0:
            if prev_vpn is not None and prev_vpn < 0:
                frac = -prev_vpn / (vpn - prev_vpn)
                return (ano - 1) + frac
            return float(ano)
        prev_vpn = vpn
    return None


def calcular_vpn(capex: float, fluxo: list[dict]) -> float:
    """VPL ao final do período analisado."""
    if fluxo:
        return fluxo[-1]["VPL acumulado (R$)"]
    return -capex


def calcular_tir_real(
    capex: float,
    geracao_ano1_kwh: float,
    tarifa_base: float,
    degradacao_anual: float = 0.006,
    escalada_tarifa: float = 0.05,
    anos: int = 25,
    opex_anual_rs: float = 0.0,
) -> Optional[float]:
    """TIR real via bissecção, considerando degradação + escalada."""
    lo, hi = -0.99, 5.0
    for _ in range(150):
        mid = (lo + hi) / 2
        npv = -capex
        for ano in range(1, anos + 1):
            geracao = geracao_ano1_kwh * ((1 - degradacao_anual) ** (ano - 1))
            tarifa = tarifa_base * ((1 + escalada_tarifa) ** (ano - 1))
            fc = geracao * tarifa - opex_anual_rs
            npv += fc / ((1 + mid) ** ano)
        if npv > 0:
            lo = mid
        else:
            hi = mid
    tir = (lo + hi) / 2
    return tir if -0.98 < tir < 4.9 else None


# ---------------------------------------------------------------------------
# CO₂ e equivalências
# ---------------------------------------------------------------------------

def calcular_co2_evitado(
    geracao_kwh_25anos: float,
    fator_emissao: float = FATOR_EMISSAO_SIN_TCO2_MWH,
) -> dict:
    co2_t = geracao_kwh_25anos / 1000.0 * fator_emissao
    arvores = co2_t * ARVORES_POR_TCO2_ANO
    return {
        "co2_evitado_t": round(co2_t, 1),
        "arvores_equivalentes": round(arvores, 0),
        "co2_evitado_kg": round(co2_t * 1000, 0),
    }


# ---------------------------------------------------------------------------
# Export Excel
# ---------------------------------------------------------------------------

def gerar_excel_completo(
    historico_df: pd.DataFrame,
    resultados_dim: dict,
    fluxo_caixa: list[dict],
    eco_metrics: dict,
    co2: dict,
    premissas: dict,
) -> bytes:
    """
    Gera arquivo Excel (.xlsx) em memória com múltiplas abas.
    Retorna bytes prontos para st.download_button.
    """
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        wb = writer.book
        fmt_header = wb.add_format({"bold": True, "bg_color": "#0f2546", "font_color": "#ffffff", "border": 1})
        fmt_num2 = wb.add_format({"num_format": '#,##0.00', "border": 1})
        fmt_num0 = wb.add_format({"num_format": '#,##0', "border": 1})
        fmt_pct = wb.add_format({"num_format": '0.00%', "border": 1})

        # Aba 1 — Premissas
        df_prem = pd.DataFrame(list(premissas.items()), columns=["Parâmetro", "Valor"])
        df_prem.to_excel(writer, sheet_name="Premissas", index=False)
        ws = writer.sheets["Premissas"]
        ws.set_column("A:A", 30)
        ws.set_column("B:B", 25)

        # Aba 2 — Histórico de Consumo
        if historico_df is not None and not historico_df.empty:
            historico_df.to_excel(writer, sheet_name="Histórico Consumo", index=False)
            ws2 = writer.sheets["Histórico Consumo"]
            ws2.set_column("A:A", 8)
            ws2.set_column("B:C", 18)

        # Aba 3 — Dimensionamento
        df_dim = pd.DataFrame([resultados_dim])
        df_dim.to_excel(writer, sheet_name="Dimensionamento", index=False)
        ws3 = writer.sheets["Dimensionamento"]
        ws3.set_column("A:Z", 22)

        # Aba 4 — Fluxo de Caixa 25 anos
        df_fc = pd.DataFrame(fluxo_caixa)
        df_fc.to_excel(writer, sheet_name="Fluxo de Caixa 25 Anos", index=False)
        ws4 = writer.sheets["Fluxo de Caixa 25 Anos"]
        ws4.set_column("A:A", 6)
        ws4.set_column("B:G", 20)
        # Adiciona gráfico de VPL acumulado
        chart = wb.add_chart({"type": "line"})
        chart.add_series({
            "name": "VPL Acumulado (R$)",
            "categories": ["Fluxo de Caixa 25 Anos", 1, 0, len(fluxo_caixa), 0],
            "values": ["Fluxo de Caixa 25 Anos", 1, 6, len(fluxo_caixa), 6],
            "line": {"color": "#2ecc71"},
        })
        chart.set_title({"name": "Evolução do VPL Acumulado"})
        chart.set_x_axis({"name": "Ano"})
        chart.set_y_axis({"name": "R$"})
        ws4.insert_chart("I2", chart, {"x_scale": 1.8, "y_scale": 1.5})

        # Aba 5 — Indicadores Econômicos
        indicadores = {
            "Economia anual ano 1 (R$)": eco_metrics.get("economia_anual", 0),
            "Payback simples (anos)": eco_metrics.get("payback", 0),
            "Payback descontado (anos)": eco_metrics.get("payback_descontado", 0),
            "VPL 25 anos (R$)": eco_metrics.get("vpn", 0),
            "TIR real (%)": (eco_metrics.get("tir_real") or 0) * 100,
            "ROI 25 anos (%)": eco_metrics.get("roi", 0),
            "CO₂ evitado 25 anos (t)": co2.get("co2_evitado_t", 0),
            "Árvores equivalentes": co2.get("arvores_equivalentes", 0),
        }
        df_ind = pd.DataFrame(list(indicadores.items()), columns=["Indicador", "Valor"])
        df_ind.to_excel(writer, sheet_name="Indicadores", index=False)
        ws5 = writer.sheets["Indicadores"]
        ws5.set_column("A:A", 32)
        ws5.set_column("B:B", 18)

    buffer.seek(0)
    return buffer.read()
