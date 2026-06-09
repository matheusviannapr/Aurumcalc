import io
import json
import os
import zipfile
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from pv_calculator import (
    aplicar_restricao_area_resultados,
    calcular_geracao_horaria_pvwatts,
    carregar_dados_equipamentos,
    get_pvwatts_last_error,
    geocode_location,
    realizar_dimensionamento_completo,
    salvar_novo_inversor,
    salvar_novo_painel,
)
from src.client_db import (
    delete_project,
    export_all_projects,
    import_projects_from_json,
    list_projects,
    load_project,
    storage_backend,
    upsert_project,
)
from src.financials import (
    calcular_co2_evitado,
    calcular_fluxo_caixa,
    calcular_payback_descontado,
    calcular_tir_real,
    calcular_vpn,
    gerar_excel_completo,
)
from src.load_profile_fitting import fit_load_profile_to_bill
from src.load_profiles import load_typical_profiles
from src.map_area import (
    calculate_azimuth_from_points,
    calculate_polygon_area_m2,
    geocode_address,
    render_area_map,
)
from src.peak_shaving import analyze_peak_shaving
from src.self_consumption import analyze_self_consumption
from src.tariff_periods import classify_peak_hours
from src.load_shifting import simulate_simplified_load_shifting

APP_BUILD = "v2-plotly-financials-2026-06-09"

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------
FILE_PATH_EQUIPAMENTOS = "BDFotovoltaica.xlsx"
MONTHS = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
STEP_NAMES = [
    "1. Identificação do Projeto",
    "2. Histórico Energético",
    "3. Perfil de Consumo",
    "4. Área Disponível",
    "5. Dimensionamento Técnico",
    "6. Análise Energética Avançada",
    "7. Viabilidade Econômica",
    "8. Relatório Final",
]
COMPLETED_STEPS_KEY = "completed_steps"

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
if "PVWATTS_API_KEY" in st.secrets:
    os.environ["PVWATTS_API_KEY"] = st.secrets["PVWATTS_API_KEY"]

st.set_page_config(
    page_title="PACE Calculator | Diagnóstico Energético",
    layout="wide",
    initial_sidebar_state="expanded",
    page_icon="⚡",
)

st.markdown(
    """
<style>
.stApp { background: linear-gradient(180deg, #061127 0%, #0b1f3a 100%); color: #eef4ff; }
section[data-testid="stSidebar"] { background: #081630; }
.card { background: #0f2546; border-radius: 14px; padding: 1rem 1.2rem; border: 1px solid #264d86; margin-bottom: 0.6rem; }
.kpi-box { background: #112244; border-radius: 12px; padding: 0.9rem 1rem; border: 1px solid #2c5fa5; text-align: center; }
.kpi-value { font-size: 1.6rem; font-weight: 700; color: #7ec8e3; }
.kpi-label { font-size: 0.78rem; color: #90a8c8; margin-top: 2px; }
.step-pill { background: #123160; color: #d9e8ff; padding: 0.35rem 0.9rem; border-radius: 999px; font-size: 0.85rem; font-weight: 600; }
.co2-badge { background: #0d3320; border: 1px solid #1a6b3a; border-radius: 10px; padding: 0.6rem 1rem; color: #4cde8a; }
hr { border-color: #1f3f70; }
.stButton > button { border-radius: 8px; }
</style>
""",
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def fmt_num(x, casas=2, default="-"):
    try:
        s = f"{float(x):,.{casas}f}"
        return s.replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return default


def kpi_card(label: str, value: str, col=None):
    html = f'<div class="kpi-box"><div class="kpi-value">{value}</div><div class="kpi-label">{label}</div></div>'
    target = col if col is not None else st
    target.markdown(html, unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def _load_equipment_cached(file_path: str):
    try:
        return carregar_dados_equipamentos(file_path)
    except Exception:
        return pd.DataFrame(), pd.DataFrame()


@st.cache_data(show_spinner=False, ttl=3600)
def _dimensionamento_cached(monthly_target, lat, lon, azimuth, tilt):
    return realizar_dimensionamento_completo(monthly_target, lat, lon, azimuth, tilt)


@st.cache_data(show_spinner=False, ttl=3600)
def _geracao_horaria_cached(potencia_kwp, lat, lon, azimuth, tilt):
    return calcular_geracao_horaria_pvwatts(
        potencia_dc_kwp=potencia_kwp,
        latitude=lat,
        longitude=lon,
        azimuth=azimuth,
        tilt=tilt,
    ) or [0.0] * 8760


def load_data():
    return _load_equipment_cached(FILE_PATH_EQUIPAMENTOS)


# ---------------------------------------------------------------------------
# Relatório LaTeX
# ---------------------------------------------------------------------------
def _tex_escape(s: str) -> str:
    """Escapa caracteres especiais do LaTeX em strings arbitrárias."""
    replacements = [
        ("\\", "\\textbackslash{}"),
        ("&", "\\&"), ("%", "\\%"), ("$", "\\$"), ("#", "\\#"),
        ("_", "\\_"), ("{", "\\{"), ("}", "\\}"), ("~", "\\textasciitilde{}"),
        ("^", "\\textasciicircum{}"),
    ]
    for old, new in replacements:
        s = s.replace(old, new)
    return s


def _tex_str(v) -> str:
    if v is None:
        return "N/D"
    return _tex_escape(str(v))


def _tex_table(rows: list[tuple], headers: list[str]) -> str:
    cols = "l" + "r" * (len(headers) - 1)
    header_row = " & ".join(f"\\textbf{{{h}}}" for h in headers) + " \\\\"
    body = "\n".join(" & ".join(_tex_str(c) for c in row) + " \\\\" for row in rows)
    return (
        "\\begin{center}\\begin{tabular}{" + cols + "}\n"
        "\\toprule\n"
        + header_row + "\n"
        "\\midrule\n"
        + body + "\n"
        "\\bottomrule\n"
        "\\end{tabular}\\end{center}\n"
    )


def _tex_pgf_bar(coords_str: str, sym_coords: list[str], ylabel: str, title: str, color: str = "blue!60") -> str:
    sym = ",".join(sym_coords)
    return (
        "\\begin{figure}[h!]\\centering\n"
        "\\begin{tikzpicture}\n"
        "\\begin{axis}[ybar,symbolic x coords={" + sym + "},xtick=data,"
        "x tick label style={rotate=45,anchor=east,font=\\small},"
        "bar width=14pt,enlarge x limits=0.15,"
        "width=0.95\\linewidth,height=7cm,"
        f"ylabel={{{ylabel}}},title={{{title}}},"
        "grid=major,grid style={dotted,gray!40},"
        f"every axis plot/.append style={{fill={color},draw={color}}}]\n"
        f"\\addplot coordinates {{{coords_str}}};\n"
        "\\end{axis}\n"
        "\\end{tikzpicture}\n"
        f"\\caption{{{title}}}\n"
        "\\end{figure}\n"
    )


def _tex_pgf_line(coords_str: str, sym_coords: list[str], ylabel: str, title: str, color: str = "blue!70") -> str:
    sym = ",".join(sym_coords)
    return (
        "\\begin{figure}[h!]\\centering\n"
        "\\begin{tikzpicture}\n"
        "\\begin{axis}[symbolic x coords={" + sym + "},xtick=data,"
        "x tick label style={rotate=45,anchor=east,font=\\tiny},"
        "width=0.95\\linewidth,height=7cm,"
        f"ylabel={{{ylabel}}},title={{{title}}},"
        "grid=major,grid style={dotted,gray!40},"
        "mark=*,mark size=1.5pt]\n"
        f"\\addplot[color={color},thick] coordinates {{{coords_str}}};\n"
        "\\addplot[dashed,red!70,thick] coordinates {{(1,0)({len(sym_coords)},0)}};\n"
        "\\end{axis}\n"
        "\\end{tikzpicture}\n"
        f"\\caption{{{title}}}\n"
        "\\end{figure}\n"
    )


def build_latex_report_zip(report: dict) -> bytes:
    files = {}

    ex = report.get("1_resumo_executivo", {}) or {}
    prem = report.get("2_premissas", {}) or {}
    historico = report.get("3_historico_consumo", []) or []
    area_info = report.get("5_area_disponivel", {}) or {}
    sistema = report.get("6_sistema_proposto", {}) or {}
    adv = report.get("7_analise_energetica", {}) or {}
    eco = report.get("8_viabilidade_economica", {}) or {}
    co2 = eco.get("co2", {}) or {}
    fc = eco.get("fluxo_caixa_25anos", []) or []
    capex_bd = eco.get("capex_breakdown", {}) or {}
    limitacoes = report.get("9_limitacoes", []) or []
    proximos = report.get("10_proximos_passos", []) or []

    # --- Seção: Histórico de consumo ---
    sec_historico = ""
    if historico:
        meses = [_tex_str(row.get("Mês", i + 1)) for i, row in enumerate(historico)]
        consumo = [float(row.get("Consumo (kWh)") or row.get("Energia ponta (kWh)", 0) or 0) for row in historico]
        custo = [float(row.get("Custo (R$)", 0) or 0) for row in historico]
        consumo_anual = sum(consumo)
        custo_anual = sum(custo)
        tarifa_media = consumo_anual / custo_anual if custo_anual > 0 else 0

        rows_hist = list(zip(meses, [f"{v:,.0f}" for v in consumo], [f"R\\$ {v:,.2f}" for v in custo]))
        sec_historico = (
            "\\section{Histórico de Consumo Energético}\n\n"
            + _tex_table(rows_hist, ["Mês", "Consumo (kWh)", "Custo (R\\$)"])
            + f"\n\\textbf{{Consumo anual total:}} {consumo_anual:,.0f} kWh \\\\\n"
            f"\\textbf{{Custo anual total:}} R\\$ {custo_anual:,.2f} \\\\\n"
            f"\\textbf{{Tarifa média implícita:}} R\\$ {tarifa_media:.4f}/kWh\n\n"
        )
        coords_hist = " ".join(f"({m},{v:.0f})" for m, v in zip(meses, consumo))
        sec_historico += _tex_pgf_bar(coords_hist, meses, "kWh", "Histórico de consumo mensal", "blue!60")

    # --- Seção: Sistema proposto ---
    pot_kwp = float(sistema.get("sistema_potencia_total_w", 0) or 0) / 1000.0
    num_paineis = sistema.get("sistema_num_total_paineis", "N/D")
    energia_anual = float(sistema.get("energia_gerada_anual_kwh", 0) or 0)
    sec_sistema = (
        "\\section{Sistema Fotovoltaico Proposto}\n\n"
        + _tex_table([
            ("Potência total do sistema", f"{pot_kwp:.2f} kWp"),
            ("Número de módulos", _tex_str(num_paineis)),
            ("Modelo do painel", _tex_str(sistema.get("painel_modelo"))),
            ("Fabricante do painel", _tex_str(sistema.get("painel_fabricante"))),
            ("Potência do painel", f"{_tex_str(sistema.get('painel_potencia'))} Wp"),
            ("Módulos em série", _tex_str(sistema.get("arranjo_modulos_serie"))),
            ("Strings paralelas/MPPT", _tex_str(sistema.get("arranjo_conjuntos_paralelo_por_mppt"))),
            ("Inversor", _tex_str(sistema.get("inversor_modelo"))),
            ("Fabricante inversor", _tex_str(sistema.get("inversor_fabricante"))),
            ("Quantidade de inversores", _tex_str(sistema.get("inversor_num_unidades"))),
            ("MPPTs utilizados", _tex_str(sistema.get("inversor_num_mppt"))),
            ("Energia gerada anual", f"{energia_anual:,.0f} kWh"),
            ("Área necessária (estimada)", f"{float(area_info.get('area_m2') or 0):.1f} m\\textsuperscript{{2}}"),
            ("Fator de aproveitamento", f"{float(area_info.get('fator_aproveitamento') or 0.7):.0%}"),
            ("Fonte do cálculo", _tex_str(sistema.get("fonte_geracao"))),
        ], ["Parâmetro", "Valor"])
    )

    # --- Seção: Análise energética ---
    sc = adv.get("self", {}) or {}
    ps = adv.get("peak", {}) or {}
    sec_energetica = ""
    if sc or ps:
        sec_energetica = (
            "\\section{Análise Energética Avançada}\n\n"
            "\\subsection{Autoconsumo e autossuficiência}\n\n"
            + _tex_table([
                ("Taxa de autoconsumo", f"{float(sc.get('self_consumption_ratio', 0) or 0)*100:.1f}\\%"),
                ("Taxa de autossuficiência", f"{float(sc.get('self_sufficiency_ratio', 0) or 0)*100:.1f}\\%"),
                ("Injeção na rede (kWh/dia)", f"{float(sc.get('total_grid_export_kwh', 0) or 0):.2f}"),
                ("Consumo da rede (kWh/dia)", f"{float(sc.get('total_grid_import_kwh', 0) or 0):.2f}"),
                ("Peak shaving (kW)", f"{float(ps.get('peak_shaving_kw', 0) or 0):.2f}"),
                ("Redução de pico (%)", f"{float(ps.get('peak_reduction_percent', 0) or 0):.1f}\\%"),
            ], ["Indicador", "Valor"])
        )

    # --- Seção: CAPEX ---
    sec_capex = ""
    if capex_bd:
        capex_rows = [(k, f"R\\$ {v:,.2f}") for k, v in capex_bd.items()]
        capex_total = eco.get("capex_total", sum(capex_bd.values()))
        capex_rows.append(("\\textbf{Total}", f"\\textbf{{R\\$ {float(capex_total):,.2f}}}"))
        sec_capex = (
            "\\section{Composição do Investimento (CAPEX)}\n\n"
            + _tex_table(capex_rows, ["Item", "Valor (R\\$)"])
        )

    # --- Seção: Indicadores financeiros ---
    economia_anual = float(eco.get("economia_anual", 0) or 0)
    payback = eco.get("payback")
    payback_desc = eco.get("payback_descontado")
    vpn = float(eco.get("vpn", 0) or 0)
    tir = eco.get("tir_real")
    roi = eco.get("roi")

    sec_financeira = (
        "\\section{Viabilidade Econômico-Financeira}\n\n"
        "\\subsection{Premissas do modelo financeiro}\n\n"
        + _tex_table([
            ("Tarifa base (R\\$/kWh)", f"R\\$ {float(eco.get('tarifa_base', 0) or 0):.4f}"),
            ("Degradação anual dos painéis", f"{float(eco.get('degradacao_anual', 0.006) or 0.006)*100:.1f}\\%/ano"),
            ("Escalada tarifária anual", f"{float(eco.get('escalada_tarifa', 0.05) or 0.05)*100:.1f}\\%/ano"),
            ("Taxa de desconto (WACC)", f"{float(eco.get('taxa_desconto', 0.10) or 0.10)*100:.1f}\\%"),
            ("OPEX anual", f"R\\$ {float(eco.get('opex_anual_rs', 0) or 0):,.2f}"),
            ("Horizonte de análise", "25 anos"),
        ], ["Premissa", "Valor"])
        + "\n\\subsection{Indicadores de retorno}\n\n"
        + _tex_table([
            ("Economia ano 1 (R\\$)", f"R\\$ {economia_anual:,.2f}"),
            ("Payback simples", f"{float(payback):.1f} anos" if payback else "N/D"),
            ("Payback descontado", f"{float(payback_desc):.1f} anos" if payback_desc else "Não atingido em 25 anos"),
            ("VPL 25 anos (R\\$)", f"R\\$ {vpn:,.2f}"),
            ("TIR real", f"{float(tir)*100:.2f}\\%" if tir is not None else "N/D"),
            ("ROI 25 anos", f"{float(roi):.1f}\\%" if roi is not None else "N/D"),
        ], ["Indicador", "Valor"])
    )

    # --- Gráfico: VPL acumulado 25 anos ---
    sec_graficos = "\\section{Análises Gráficas}\n\n"
    if historico:
        coords_hist = " ".join(f"({m},{v:.0f})" for m, v in zip(meses, consumo))
        sec_graficos += _tex_pgf_bar(coords_hist, meses, "kWh", "Histórico de consumo mensal", "blue!60")

    if fc:
        anos_fc = [str(r["Ano"]) for r in fc]
        vpls_fc = [float(r.get("VPL acumulado (R$)", 0) or 0) for r in fc]
        econ_fc = [float(r.get("Fluxo líquido (R$)", 0) or 0) for r in fc]
        coords_vpl = " ".join(f"({a},{v:.0f})" for a, v in zip(anos_fc, vpls_fc))
        coords_fc2 = " ".join(f"({a},{v:.0f})" for a, v in zip(anos_fc, econ_fc))
        sec_graficos += _tex_pgf_line(coords_vpl, anos_fc, "R\\$", "Evolução do VPL acumulado (25 anos)")
        sec_graficos += _tex_pgf_bar(coords_fc2, anos_fc, "R\\$", "Fluxo de caixa líquido anual", "green!50!black!70")

    sens = eco.get("sensibilidade_tarifaria", {}) or {}
    if sens:
        s_labels = list(sens.keys())
        s_values = [float(v or 0) for v in sens.values()]
        coords_sens = " ".join(f"({l},{v:.0f})" for l, v in zip(s_labels, s_values))
        sec_graficos += _tex_pgf_bar(coords_sens, s_labels, "R\\$", "Sensibilidade da economia à tarifa", "orange!70")

    # --- Seção: CO2 ---
    sec_co2 = ""
    if co2:
        sec_co2 = (
            "\\section{Impacto Ambiental}\n\n"
            + _tex_table([
                ("CO\\textsubscript{2} evitado em 25 anos", f"{float(co2.get('co2_evitado_t', 0)):,.1f} t"),
                ("CO\\textsubscript{2} evitado (kg)", f"{float(co2.get('co2_evitado_kg', 0)):,.0f} kg"),
                ("Equivalência em árvores plantadas", f"{float(co2.get('arvores_equivalentes', 0)):,.0f} árvores"),
                ("Fator de emissão SIN (MCTIC 2023)", "0,0839 tCO\\textsubscript{2}eq/MWh"),
            ], ["Indicador ambiental", "Valor"])
            + "\n\\textit{Metodologia: fator de emissão da rede SIN conforme MCTIC 2023. "
            "Equivalência de árvores baseada em absorção média de 100 kg CO\\textsubscript{2}/árvore/ano.}\n\n"
        )

    # --- Fluxo de caixa completo ---
    sec_fluxo = ""
    if fc:
        fc_rows = [
            (
                str(r["Ano"]),
                f"{float(r.get('Geração (kWh)', 0)):,.0f}",
                f"R\\$ {float(r.get('Tarifa (R$/kWh)', 0)):.4f}",
                f"R\\$ {float(r.get('Fluxo líquido (R$)', 0)):,.2f}",
                f"R\\$ {float(r.get('Fluxo descontado (R$)', 0)):,.2f}",
                f"R\\$ {float(r.get('VPL acumulado (R$)', 0)):,.2f}",
            )
            for r in fc
        ]
        sec_fluxo = (
            "\\section{Fluxo de Caixa — 25 Anos}\n\n"
            "\\begin{center}\\footnotesize\n"
            "\\begin{tabular}{rrrrrrr}\n"
            "\\toprule\n"
            "\\textbf{Ano} & \\textbf{Geração (kWh)} & \\textbf{Tarifa} & "
            "\\textbf{Fluxo líq.} & \\textbf{Fluxo desc.} & \\textbf{VPL acum.} \\\\\n"
            "\\midrule\n"
            + "\n".join(" & ".join(r) + " \\\\" for r in fc_rows)
            + "\n\\bottomrule\n\\end{tabular}\n\\end{center}\n\n"
        )

    # --- Limitações e próximos passos ---
    sec_limite = ""
    if limitacoes:
        items = "\n".join(f"\\item {_tex_str(l)}" for l in limitacoes)
        sec_limite = f"\\section{{Limitações e Alertas}}\n\\begin{{itemize}}\n{items}\n\\end{{itemize}}\n\n"

    prox_items = "\n".join(f"\\item {_tex_str(p)}" for p in (proximos or [
        "Validar curva com medição real de carga (15 min).",
        "Refinar premissas tarifárias da distribuidora local.",
        "Realizar visita técnica para engenharia de instalação.",
        "Contratar laudos de engenharia (ART/RRT).",
        "Solicitar aprovação à distribuidora (acesso à rede).",
    ]))

    latex = (
        "\\documentclass[11pt,a4paper]{article}\n"
        "\\usepackage[utf8]{inputenc}\n"
        "\\usepackage[T1]{fontenc}\n"
        "\\usepackage[brazil]{babel}\n"
        "\\usepackage{graphicx}\n"
        "\\usepackage{booktabs}\n"
        "\\usepackage{geometry}\n"
        "\\usepackage{pgfplots}\n"
        "\\usepackage{xcolor}\n"
        "\\usepackage{titlesec}\n"
        "\\usepackage{fancyhdr}\n"
        "\\usepackage{hyperref}\n"
        "\\pgfplotsset{compat=1.18}\n"
        "\\geometry{margin=2.2cm}\n"
        "\\hypersetup{colorlinks=true,linkcolor=blue!60!black,urlcolor=blue!60!black}\n"
        "\\pagestyle{fancy}\n"
        "\\fancyhf{}\n"
        "\\fancyhead[L]{\\textbf{PACE Calculator} --- Relatório Técnico Fotovoltaico}\n"
        "\\fancyhead[R]{\\today}\n"
        "\\fancyfoot[C]{\\thepage}\n"
        "\\titleformat{\\section}{\\large\\bfseries\\color{blue!40!black}}{\\thesection}{1em}{}\n"
        "\\titleformat{\\subsection}{\\normalsize\\bfseries}{\\thesubsection}{1em}{}\n"
        "\\begin{document}\n\n"
        "\\begin{titlepage}\n"
        "\\centering\n"
        "{\\Huge\\bfseries\\color{blue!40!black} PACE Calculator\\par}\n"
        "\\vspace{0.5cm}\n"
        "{\\Large Relatório Técnico de Viabilidade\\par}\n"
        "{\\large Sistema Fotovoltaico Conectado à Rede\\par}\n"
        "\\vspace{1.5cm}\n"
        "\\begin{tabular}{ll}\n"
        f"\\textbf{{Cliente:}} & {_tex_str(ex.get('cliente'))} \\\\\n"
        f"\\textbf{{Localização:}} & {_tex_str(ex.get('local'))} \\\\\n"
        f"\\textbf{{Distribuidora:}} & {_tex_str(ex.get('distribuidora'))} \\\\\n"
        f"\\textbf{{Contexto:}} & {_tex_str(ex.get('contexto'))} \\\\\n"
        f"\\textbf{{Coordenadas:}} & Lat {_tex_str(prem.get('latitude'))}, Lon {_tex_str(prem.get('longitude'))} \\\\\n"
        f"\\textbf{{Modalidade tarifária:}} & {_tex_str(prem.get('modalidade'))} \\\\\n"
        "\\textbf{Data:} & \\today \\\\\n"
        "\\end{tabular}\n"
        "\\vfill\n"
        "{\\small Gerado automaticamente pelo \\textbf{PACE Calculator} --- Plataforma de Diagnóstico Energético Solar}\n"
        "\\end{titlepage}\n\n"
        "\\tableofcontents\n"
        "\\newpage\n\n"
        "\\section{Resumo Executivo}\n\n"
        f"Este relatório apresenta a análise técnica e econômica para implantação de um sistema "
        f"fotovoltaico conectado à rede para o cliente \\textbf{{{_tex_str(ex.get('cliente', 'N/D'))}}}, "
        f"localizado em \\textbf{{{_tex_str(ex.get('local', 'N/D'))}}}, "
        f"distribuidora \\textbf{{{_tex_str(ex.get('distribuidora', 'N/D'))}}}.\n\n"
        f"O sistema proposto possui potência de \\textbf{{{pot_kwp:.2f} kWp}}, "
        f"com geração estimada de \\textbf{{{energia_anual:,.0f} kWh/ano}} e "
        f"economia de \\textbf{{R\\$ {economia_anual:,.2f}/ano}} no primeiro ano.\n\n"
        + _tex_table([
            ("Potência instalada", f"{pot_kwp:.2f} kWp"),
            ("Energia gerada anual", f"{energia_anual:,.0f} kWh"),
            ("Economia ano 1", f"R\\$ {economia_anual:,.2f}"),
            ("CAPEX total", f"R\\$ {float(eco.get('capex_total', 0) or 0):,.2f}"),
            ("Payback simples", f"{float(payback):.1f} anos" if payback else "N/D"),
            ("VPL 25 anos", f"R\\$ {vpn:,.2f}"),
            ("TIR real", f"{float(tir)*100:.2f}\\%" if tir is not None else "N/D"),
            ("CO\\textsubscript{2} evitado", f"{float(co2.get('co2_evitado_t', 0)):,.1f} t em 25 anos"),
        ], ["Indicador", "Valor"])
        + "\n" + sec_historico
        + "\n" + sec_sistema
        + "\n" + sec_energetica
        + "\n" + sec_capex
        + "\n" + sec_financeira
        + "\n" + sec_graficos
        + "\n" + sec_co2
        + "\n" + sec_fluxo
        + "\n" + sec_limite
        + "\\section{Próximos Passos}\n\n"
        f"\\begin{{itemize}}\n{prox_items}\n\\end{{itemize}}\n\n"
        "\\section{Declaração e Responsabilidade}\n\n"
        "Este relatório foi gerado automaticamente pela plataforma \\textbf{PACE Calculator} com base nos dados "
        "fornecidos pelo usuário e nas estimativas de irradiação solar fornecidas pela API PVWatts (NREL). "
        "Os valores são estimativas para fins de pré-viabilidade e devem ser validados por engenheiro "
        "eletricista habilitado antes de qualquer decisão de investimento. "
        "O fator de emissão utilizado é o do Sistema Interligado Nacional (SIN), conforme publicação "
        "do MCTIC de 2023 (0,0839 tCO\\textsubscript{2}eq/MWh).\n\n"
        "\\end{document}\n"
    )

    files["relatorio.tex"] = latex.encode("utf-8")
    files["relatorio_dados.json"] = json.dumps(report, ensure_ascii=False, indent=2, default=str).encode("utf-8")
    files["README.txt"] = (
        "Como compilar o relatório:\n"
        "1. Instale uma distribuição LaTeX (TeX Live, MiKTeX)\n"
        "2. Execute: pdflatex relatorio.tex\n"
        "3. Execute novamente para gerar índice: pdflatex relatorio.tex\n"
        "Ou use Overleaf (overleaf.com) fazendo upload deste zip.\n"
    ).encode("utf-8")

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, data in files.items():
            zf.writestr(fname, data)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()


# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------
def ensure_state():
    defaults = {
        "current_step": 1,
        COMPLETED_STEPS_KEY: set(),
        "latitude": -23.5505,
        "longitude": -46.6333,
        "tariff_group": "B",
        "historico_df": pd.DataFrame({"Mês": MONTHS, "Consumo (kWh)": [None] * 12, "Custo (R$)": [None] * 12}),
        "fit_adv": None,
        "fit_adv_profile": None,
        "available_area_m2": None,
        "resultados_df_adv": None,
        "area_meta": None,
        "advanced_metrics": None,
        "economic_metrics": None,
        "active_project": None,
        "project_name": "",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


ensure_state()


def _serialize_state():
    payload = {}
    for key, value in st.session_state.items():
        if key.startswith("FormSubmitter"):
            continue
        if isinstance(value, pd.DataFrame):
            payload[key] = {"__type__": "dataframe", "records": value.to_dict(orient="records")}
        elif isinstance(value, set):
            payload[key] = {"__type__": "set", "values": list(value)}
        else:
            payload[key] = value
    return payload


def _restore_state(payload):
    for key, value in payload.items():
        if isinstance(value, dict) and value.get("__type__") == "dataframe":
            st.session_state[key] = pd.DataFrame(value.get("records", []))
        elif isinstance(value, dict) and value.get("__type__") == "set":
            st.session_state[key] = set(value.get("values", []))
        else:
            st.session_state[key] = value


def mark_step_complete(step: int):
    st.session_state[COMPLETED_STEPS_KEY].add(step)


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
st.sidebar.title("PACE Calculator • Jornada Guiada")
progress = (st.session_state["current_step"] - 1) / (len(STEP_NAMES) - 1)
st.sidebar.progress(progress, text=f"Etapa {st.session_state['current_step']} / {len(STEP_NAMES)}")

st.sidebar.markdown("**Navegação rápida:**")
completed = st.session_state.get(COMPLETED_STEPS_KEY, set())
for i, step_name in enumerate(STEP_NAMES, start=1):
    is_current = i == st.session_state["current_step"]
    is_done = i in completed
    icon = "✅" if is_done else ("▶️" if is_current else "○")
    label = f"{icon} {step_name}"
    if is_done or is_current or i <= max(completed, default=0) + 1:
        if st.sidebar.button(label, key=f"nav_{i}", use_container_width=True):
            st.session_state["current_step"] = i
            st.rerun()
    else:
        st.sidebar.markdown(f"<span style='color:#4a6080;font-size:0.85rem'>{icon} {step_name}</span>", unsafe_allow_html=True)

with st.sidebar.expander("📦 Base de equipamentos"):
    df_p, df_i = load_data()
    st.caption(f"Painéis cadastrados: {len(df_p)}")
    st.caption(f"Inversores cadastrados: {len(df_i)}")

with st.sidebar.expander("💾 Banco de projetos", expanded=False):
    try:
        backend = storage_backend()
        st.caption(f"Backend: **{backend}**")
        if backend == "SQLite (local)":
            st.caption("⚠️ SQLite não persiste no Streamlit Cloud. Configure `DATABASE_URL` nos secrets para persistência permanente.")
    except Exception:
        pass

    db_client = st.text_input("Cliente", value=st.session_state.get("cliente_nome", ""), key="sb_db_client")
    if db_client and not st.session_state.get("cliente_nome"):
        st.session_state["cliente_nome"] = db_client

    project_rows = list_projects(db_client if db_client else None)
    project_labels = [f"{row[0]} | {row[1]} | {row[2][:19]}" for row in project_rows]
    selected_label = st.selectbox("Projetos salvos", options=["Selecione..."] + project_labels, key="sb_proj_select")

    col_load, col_del = st.columns(2)
    with col_load:
        if st.button("Carregar", key="sb_load", use_container_width=True):
            if selected_label == "Selecione...":
                st.warning("Selecione um projeto.")
            else:
                idx = project_labels.index(selected_label)
                client_name, project_name, _ = project_rows[idx]
                payload = load_project(client_name, project_name)
                if payload:
                    _restore_state(payload)
                    st.session_state["active_project"] = f"{client_name}:{project_name}"
                    st.success(f"'{project_name}' carregado.")
                    st.rerun()
                else:
                    st.error("Projeto não encontrado.")
    with col_del:
        if st.button("Excluir", key="sb_del", use_container_width=True, type="secondary"):
            if selected_label == "Selecione...":
                st.warning("Selecione um projeto.")
            else:
                idx = project_labels.index(selected_label)
                client_name, project_name, _ = project_rows[idx]
                delete_project(client_name, project_name)
                st.success(f"'{project_name}' excluído.")
                st.rerun()

    st.session_state["project_name"] = st.text_input(
        "Nome do projeto",
        value=st.session_state.get("project_name", "") or st.session_state.get("localizacao", ""),
        key="sb_proj_name",
    )
    if st.button("Salvar/Atualizar", key="sb_save", use_container_width=True, type="primary"):
        client_name = (db_client or st.session_state.get("cliente_nome", "")).strip()
        project_name = st.session_state.get("project_name", "").strip()
        if not client_name or not project_name:
            st.error("Informe cliente e nome do projeto.")
        else:
            upsert_project(client_name, project_name, _serialize_state())
            st.session_state["active_project"] = f"{client_name}:{project_name}"
            st.success(f"Projeto '{project_name}' salvo.")

    st.markdown("---")
    st.caption("**Backup / Migração**")
    try:
        export_bytes = export_all_projects()
        st.download_button(
            "⬇️ Exportar todos os projetos",
            data=export_bytes,
            file_name="pace_projetos_backup.json",
            mime="application/json",
            use_container_width=True,
            key="sb_export",
        )
    except Exception as e:
        st.caption(f"Exportação indisponível: {e}")

    uploaded = st.file_uploader("⬆️ Importar backup (.json)", type="json", key="sb_import_file")
    if uploaded and st.button("Importar agora", key="sb_import_btn", use_container_width=True):
        try:
            count = import_projects_from_json(uploaded.read())
            st.success(f"{count} projetos importados.")
            st.rerun()
        except Exception as e:
            st.error(f"Falha na importação: {e}")

with st.sidebar.expander("➕ Novo equipamento"):
    tab_p, tab_i = st.tabs(["Painel", "Inversor"])
    with tab_p:
        with st.form("novo_painel", clear_on_submit=True):
            st.markdown("**Identificação**")
            modelo_p = st.text_input("Modelo *", help="Ex: Canadian Solar CS6W-550MS")
            fabricante_p = st.text_input("Fabricante *", help="Ex: Canadian Solar")
            st.markdown("**Dados elétricos (STC)**")
            pmax = st.number_input("Potência máxima Pmax (Wp) *", min_value=1.0, value=550.0, step=5.0,
                                   help="Potência nominal em condições padrão de teste (STC).")
            vmp = st.number_input("Tensão de operação ótima Vmp (V)", min_value=0.0, value=41.8, step=0.1,
                                  help="Tensão no ponto de máxima potência.")
            imp = st.number_input("Corrente de operação ótima Imp (A)", min_value=0.0, value=13.16, step=0.01,
                                  help="Corrente no ponto de máxima potência.")
            voc = st.number_input("Tensão de circuito aberto Voc (V) *", min_value=1.0, value=49.5, step=0.1,
                                  help="Tensão em circuito aberto (sem carga).")
            isc = st.number_input("Corrente de curto-circuito Isc (A) *", min_value=0.1, value=13.93, step=0.01,
                                  help="Corrente em curto-circuito.")
            eficiencia = st.number_input("Eficiência do módulo (%)", min_value=0.0, max_value=30.0, value=21.3, step=0.1,
                                         help="Eficiência de conversão do módulo fotovoltaico.")
            if st.form_submit_button("Salvar painel", type="primary"):
                if not modelo_p or not fabricante_p:
                    st.error("Modelo e fabricante são obrigatórios.")
                else:
                    ok = salvar_novo_painel({
                        "modelo": modelo_p,
                        "fabricante": fabricante_p,
                        "potencia_maxima_nominal_pmax": pmax,
                        "tensao_operacao_otima_vmp": vmp,
                        "corrente_operacao_otima_imp": imp,
                        "tensao_circuito_aberto_voc": voc,
                        "corrente_curto_circuito_isc": isc,
                        "eficiencia_modulo": eficiencia,
                    })
                    st.success("Painel salvo.") if ok else st.error("Falha ao salvar.")
                    _load_equipment_cached.clear()
    with tab_i:
        with st.form("novo_inversor", clear_on_submit=True):
            st.markdown("**Identificação**")
            modelo_i = st.text_input("Modelo *", help="Ex: Fronius Symo 15.0-3-M")
            fabricante_i = st.text_input("Fabricante *", help="Ex: Fronius")
            st.markdown("**Lado CC (entrada FV)**")
            pot_fv_max = st.number_input("Potência máx. FV entrada (W) *", min_value=1.0, value=50000.0, step=500.0,
                                         help="Máxima potência CC aceita pelo inversor.")
            vmax_cc = st.number_input("Tensão máxima CC (V) *", min_value=1.0, value=1100.0, step=10.0,
                                      help="Tensão máxima de entrada CC (string).")
            tensao_start = st.number_input("Tensão de partida (V)", min_value=1.0, value=200.0, step=10.0,
                                           help="Tensão mínima para o inversor iniciar operação.")
            tensao_nominal_cc = st.number_input("Tensão nominal CC (V)", min_value=1.0, value=600.0, step=10.0)
            faixa_mpp = st.text_input("Faixa de tensão MPP (V)", value="200-850",
                                      help="Faixa de operação do rastreador MPPT. Ex: 200-850")
            num_mppt = st.number_input("Número de MPPT trackers *", min_value=1, max_value=20, value=4, step=1)
            imax_mppt = st.number_input("Corrente máx. entrada por MPPT (A)", min_value=0.1, value=30.0, step=0.5)
            isc_mppt = st.number_input("Corrente máx. curto-circuito por MPPT (A)", min_value=0.1, value=40.0, step=0.5)
            st.markdown("**Lado CA (saída)**")
            pot_nom_ca = st.number_input("Potência nominal CA (W) *", min_value=1.0, value=50000.0, step=500.0)
            pot_ap_ca = st.number_input("Potência máx. aparente CA (VA)", min_value=1.0, value=50000.0, step=500.0)
            vnom_ca = st.number_input("Tensão nominal CA (V)", min_value=100.0, value=380.0, step=10.0)
            freq_ca = st.selectbox("Frequência da rede", ["60Hz", "50Hz"])
            imax_saida = st.number_input("Corrente de saída máxima (A)", min_value=0.1, value=80.0, step=1.0)
            fp_ajust = st.text_input("Fator de potência ajustável", value="0.8i-0.8c",
                                     help="Faixa de ajuste do fator de potência. Ex: 0.8i-0.8c")
            fases_ca = st.selectbox("Fases CA", [1, 2, 3], index=2)
            if st.form_submit_button("Salvar inversor", type="primary"):
                if not modelo_i or not fabricante_i:
                    st.error("Modelo e fabricante são obrigatórios.")
                else:
                    ok = salvar_novo_inversor({
                        "modelo": modelo_i,
                        "fabricante": fabricante_i,
                        "potencia_maxima_fv_maxima": pot_fv_max,
                        "tensao_maxima_cc": vmax_cc,
                        "tensao_start": tensao_start,
                        "tensao_nominal": tensao_nominal_cc,
                        "faixa_tensao_mpp": faixa_mpp,
                        "numero_mpp_trackers": int(num_mppt),
                        "corrente_maxima_entrada_por_mppt_tracker": imax_mppt,
                        "corrente_maxima_curto_circuito_por_mppt_tracker": isc_mppt,
                        "maxima_potencia_nominal_ca": pot_nom_ca,
                        "potencia_maxima_aparente_ca": pot_ap_ca,
                        "tensao_nominal_ca": vnom_ca,
                        "frequencia_rede_ca": freq_ca,
                        "corrente_saida_maxima": imax_saida,
                        "fator_potencia_ajustavel": fp_ajust,
                        "quantidade_fases_ca": int(fases_ca),
                    })
                    st.success("Inversor salvo.") if ok else st.error("Falha ao salvar.")
                    _load_equipment_cached.clear()

# ---------------------------------------------------------------------------
# Header principal
# ---------------------------------------------------------------------------
st.title("⚡ PACE Calculator — Plataforma de Diagnóstico Energético Solar")
st.caption("Do diagnóstico ao relatório final: jornada consultiva, guiada e estratégica.")

step = st.session_state["current_step"]
st.markdown(f"<span class='step-pill'>Etapa atual: {STEP_NAMES[step-1]}</span>", unsafe_allow_html=True)
st.markdown("---")

# ===========================================================================
# ETAPA 1 — Identificação do Projeto
# ===========================================================================
if step == 1:
    st.header("ETAPA 1 — Identificação do Projeto")
    st.info("Vamos primeiro entender quem é o consumidor e qual é o contexto energético do projeto.")
    c1, c2 = st.columns(2)
    with c1:
        st.session_state["cliente_nome"] = st.text_input(
            "Nome do cliente",
            value=st.session_state.get("cliente_nome", ""),
            help="Nome completo ou razão social do cliente.",
        )
        st.session_state["cliente_tipo"] = st.selectbox(
            "Tipo de cliente",
            ["Comercial", "Industrial", "Rural", "Residencial", "Poder Público"],
            help="Segmento do consumidor. Afeta o perfil de carga típico sugerido na Etapa 3.",
        )
        st.session_state["tariff_group"] = st.radio(
            "Grupo tarifário",
            ["A", "B"],
            horizontal=True,
            help="Grupo A: alta/média tensão (demanda contratada). Grupo B: baixa tensão (apenas energia).",
        )
        st.session_state["modalidade"] = st.selectbox(
            "Modalidade tarifária",
            ["Convencional", "Branca", "Verde", "Azul"],
            help="Convencional/Branca: Grupo B. Verde/Azul: Grupo A.",
        )
    with c2:
        st.session_state["localizacao"] = st.text_input(
            "Localização",
            value=st.session_state.get("localizacao", "São Paulo, SP"),
            help="Cidade/estado ou endereço. Usado para geocodificar coordenadas automaticamente.",
        )
        if st.button("📍 Buscar coordenadas", help="Geocodifica o endereço via OpenStreetMap."):
            with st.spinner("Buscando coordenadas..."):
                lat_geo, lon_geo = geocode_location(st.session_state["localizacao"])
            if lat_geo is not None:
                st.session_state["latitude"] = float(lat_geo)
                st.session_state["longitude"] = float(lon_geo)
                st.rerun()
            else:
                st.error("Não foi possível localizar o endereço.")
        st.session_state["latitude"] = st.number_input(
            "Latitude",
            value=float(st.session_state["latitude"]),
            format="%.6f",
            help="Latitude decimal (negativo = sul). Essencial para o cálculo solar.",
        )
        st.session_state["longitude"] = st.number_input(
            "Longitude",
            value=float(st.session_state["longitude"]),
            format="%.6f",
            help="Longitude decimal (negativo = oeste). Essencial para o cálculo solar.",
        )
        st.session_state["distribuidora"] = st.text_input(
            "Distribuidora",
            value=st.session_state.get("distribuidora", ""),
            help="Nome da concessionária local (ex: ENEL SP, CEMIG, CPFL).",
        )
    st.session_state["observacoes"] = st.text_area(
        "Observações do projeto",
        value=st.session_state.get("observacoes", ""),
    )

    if st.session_state.get("cliente_nome") and st.session_state.get("localizacao"):
        mark_step_complete(1)

# ===========================================================================
# ETAPA 2 — Histórico Energético
# ===========================================================================
elif step == 2:
    st.header("ETAPA 2 — Histórico Energético")
    st.info("Esta etapa constrói a base energética anual do cliente.")

    group = st.session_state["tariff_group"]
    if group == "B":
        base_df = st.session_state.get("historico_df")
        hist_df = st.data_editor(base_df, use_container_width=True, num_rows="fixed")
        st.session_state["historico_df"] = hist_df

        if st.button("Preencher meses faltantes pela média"):
            for col in ["Consumo (kWh)", "Custo (R$)"]:
                series = pd.to_numeric(hist_df[col], errors="coerce")
                avg = float(series.dropna().mean()) if not series.dropna().empty else 0.0
                hist_df[col] = series.fillna(avg)
            st.session_state["historico_df"] = hist_df
            st.rerun()

        real_count = int(pd.to_numeric(hist_df["Consumo (kWh)"], errors="coerce").notna().sum())
        annual_kwh = float(pd.to_numeric(hist_df["Consumo (kWh)"], errors="coerce").fillna(0).sum())
        annual_cost = float(pd.to_numeric(hist_df["Custo (R$)"], errors="coerce").fillna(0).sum())

        c1, c2, c3 = st.columns(3)
        c1.metric("Consumo anual (kWh)", fmt_num(annual_kwh, 0))
        c2.metric("Custo anual (R$)", fmt_num(annual_cost, 2))
        c3.metric("Meses reais informados", f"{real_count} / 12")

        if annual_kwh > 0:
            kwh_vals = pd.to_numeric(hist_df["Consumo (kWh)"], errors="coerce").fillna(0).tolist()
            fig = px.bar(
                x=MONTHS,
                y=kwh_vals,
                labels={"x": "Mês", "y": "Consumo (kWh)"},
                title="Histórico de consumo mensal",
                color_discrete_sequence=["#2c7be5"],
            )
            fig.update_layout(
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
                font_color="#eef4ff",
            )
            st.plotly_chart(fig, use_container_width=True)
            mark_step_complete(2)
    else:
        if "historico_df_a" not in st.session_state:
            st.session_state["historico_df_a"] = pd.DataFrame(
                {
                    "Mês": MONTHS,
                    "Energia ponta (kWh)": [None] * 12,
                    "Energia fora ponta (kWh)": [None] * 12,
                    "Demanda ponta (kW)": [None] * 12,
                    "Demanda fora ponta (kW)": [None] * 12,
                    "Custo (R$)": [None] * 12,
                }
            )
        hist_df_a = st.data_editor(st.session_state["historico_df_a"], use_container_width=True, num_rows="fixed")
        st.session_state["historico_df_a"] = hist_df_a

        if st.button("Replicar meses faltantes por média (Grupo A)"):
            for col in hist_df_a.columns[1:]:
                series = pd.to_numeric(hist_df_a[col], errors="coerce")
                avg = float(series.dropna().mean()) if not series.dropna().empty else 0.0
                hist_df_a[col] = series.fillna(avg)
            st.session_state["historico_df_a"] = hist_df_a
            st.rerun()

        ep_vals = pd.to_numeric(hist_df_a["Energia ponta (kWh)"], errors="coerce").fillna(0)
        ef_vals = pd.to_numeric(hist_df_a["Energia fora ponta (kWh)"], errors="coerce").fillna(0)
        energia_total = float(ep_vals.sum() + ef_vals.sum())
        st.metric("Energia anual total (kWh)", fmt_num(energia_total, 0))

        if energia_total > 0:
            fig = go.Figure()
            fig.add_bar(x=MONTHS, y=ep_vals.tolist(), name="Ponta", marker_color="#e74c3c")
            fig.add_bar(x=MONTHS, y=ef_vals.tolist(), name="Fora ponta", marker_color="#2ecc71")
            fig.update_layout(
                barmode="stack",
                title="Energia mensal por posto tarifário",
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
                font_color="#eef4ff",
                xaxis_title="Mês",
                yaxis_title="kWh",
            )
            st.plotly_chart(fig, use_container_width=True)
            mark_step_complete(2)

# ===========================================================================
# ETAPA 3 — Perfil de Consumo
# ===========================================================================
elif step == 3:
    st.header("ETAPA 3 — Perfil de Consumo")
    st.info("Agora traduzimos sua conta de energia em comportamento energético ao longo do dia.")
    profiles = load_typical_profiles()
    profile_map = {p["nome"]: p for p in profiles}

    tipo = st.selectbox(
        "Perfil típico de consumo",
        options=list(profile_map.keys()),
        help="Curva sintética representativa do segmento. Será calibrada com os dados da conta.",
    )
    selected = profile_map[tipo]
    base_curve = selected["curva_24h_pu"]

    st.markdown("#### Curva característica (24h)")
    use_custom_curve = st.checkbox(
        "Editar curva horária manualmente antes do fit",
        value=False,
        help="Permite ajustar os pesos hora a hora antes da calibração automática.",
    )
    if "custom_curve_24h" not in st.session_state or not isinstance(st.session_state.get("custom_curve_24h"), list):
        st.session_state["custom_curve_24h"] = list(base_curve)
    if st.button("Resetar curva para perfil típico"):
        st.session_state["custom_curve_24h"] = list(base_curve)

    if use_custom_curve:
        custom_df = pd.DataFrame({"Hora": list(range(24)), "Peso relativo (pu)": st.session_state["custom_curve_24h"]})
        edited = st.data_editor(
            custom_df,
            use_container_width=True,
            num_rows="fixed",
            column_config={
                "Hora": st.column_config.NumberColumn("Hora", disabled=True),
                "Peso relativo (pu)": st.column_config.NumberColumn("Peso (pu)", min_value=0.0, step=0.01, format="%.4f"),
            },
            key="curve_editor_24h",
        )
        edited_vals = pd.to_numeric(edited["Peso relativo (pu)"], errors="coerce").fillna(0.0).clip(lower=0.0).tolist()
        if len(edited_vals) == 24:
            soma = float(sum(edited_vals))
            if soma > 0:
                st.session_state["custom_curve_24h"] = [v / soma for v in edited_vals]
                st.caption("Curva normalizada para soma = 1,00.")
            else:
                st.warning("Soma zerada. Ajuste ao menos uma hora.")
    else:
        st.session_state["custom_curve_24h"] = list(base_curve)

    curve_to_show = st.session_state["custom_curve_24h"] if use_custom_curve else base_curve
    fig_curve = px.area(
        x=list(range(24)),
        y=curve_to_show,
        labels={"x": "Hora", "y": "Peso relativo (pu)"},
        title="Perfil de consumo horário",
        color_discrete_sequence=["#2c7be5"],
    )
    fig_curve.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font_color="#eef4ff",
    )
    st.plotly_chart(fig_curve, use_container_width=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        dias_func = st.number_input(
            "Dias de funcionamento/mês",
            min_value=1, max_value=31, value=22,
            help="Número de dias que o estabelecimento opera por mês.",
        )
    with c2:
        peak_start = st.text_input("Início ponta", value="18:00", help="Horário de início do período de ponta (HH:MM).")
    with c3:
        peak_end = st.text_input("Fim ponta", value="21:00", help="Horário de fim do período de ponta (HH:MM).")

    group = st.session_state["tariff_group"]
    st.markdown("#### Dados da conta para calibração")
    if group == "A":
        c1, c2 = st.columns(2)
        with c1:
            ep = st.number_input("Energia ponta mensal (kWh)", min_value=0.0, value=1000.0)
            dp = st.number_input("Demanda ponta (kW)", min_value=0.0, value=120.0)
        with c2:
            ef = st.number_input("Energia fora ponta mensal (kWh)", min_value=0.0, value=4000.0)
            dfp = st.number_input("Demanda fora ponta (kW)", min_value=0.0, value=90.0)
        bill = {"energia_ponta_kwh_mes": ep, "energia_fora_ponta_kwh_mes": ef, "demanda_ponta_kw": dp, "demanda_fora_ponta_kw": dfp}
    else:
        energy_total = st.number_input(
            "Energia total mensal (kWh)",
            min_value=0.0, value=5000.0,
            help="Energia total consumida por mês conforme conta de luz.",
        )
        bill = {"energia_total_kwh_mes": energy_total, "demanda_estimada_kw": 0.0}

    if st.button("▶ Calibrar curva e gerar fit", type="primary"):
        peak_labels = classify_peak_hours(list(range(24)), peak_start, peak_end, weekdays_only=False)
        curve_for_fit = st.session_state["custom_curve_24h"] if use_custom_curve else selected["curva_24h_pu"]
        fit = fit_load_profile_to_bill(
            base_profile_24h=curve_for_fit,
            bill_data=bill,
            operation_days=int(dias_func),
            peak_hours=peak_labels,
            group_type=group,
        )
        st.session_state["fit_adv"] = fit
        st.session_state["fit_adv_profile"] = selected
        st.success("Curva calibrada com sucesso.")
        mark_step_complete(3)

    fit_adv = st.session_state.get("fit_adv")
    if fit_adv is not None:
        sel = st.session_state["fit_adv_profile"]
        horas = list(range(24))
        fig_fit = go.Figure()
        fig_fit.add_scatter(x=horas, y=sel["curva_24h_pu"], name="Perfil típico (pu)", line={"dash": "dot", "color": "#7ec8e3"})
        fig_fit.add_scatter(x=horas, y=fit_adv["load_curve_kw"], name="Curva ajustada (kW)", fill="tozeroy", line={"color": "#2ecc71"})
        fig_fit.update_layout(
            title="Curva de carga calibrada",
            xaxis_title="Hora",
            yaxis_title="kW",
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font_color="#eef4ff",
        )
        st.plotly_chart(fig_fit, use_container_width=True)
        warns = fit_adv.get("warnings", [])
        for w in warns:
            st.warning(w)
        if not warns:
            st.success("Fit sem inconsistências críticas.")

# ===========================================================================
# ETAPA 4 — Área Disponível
# ===========================================================================
elif step == 4:
    st.header("ETAPA 4 — Área Disponível")
    st.info("Aqui definimos a limitação física real do projeto.")
    endereco = st.text_input("1) Endereço para buscar no mapa", value=st.session_state.get("localizacao", "São Paulo, SP"))

    c1, c2 = st.columns(2)
    with c1:
        if st.button("2) Geocodificar e centralizar mapa"):
            with st.spinner("Geocodificando..."):
                try:
                    lat_map, lon_map = geocode_address(endereco)
                    st.session_state["adv_map_center"] = (lat_map, lon_map)
                    st.success(f"Mapa centralizado em {lat_map:.5f}, {lon_map:.5f}")
                except Exception as e:
                    st.error(f"Falha ao geocodificar: {e}")
    with c2:
        pack = st.slider(
            "5) Fator de aproveitamento da área",
            0.50, 0.90, 0.70, 0.01,
            help="Fração da área bruta efetivamente utilizável (desconta passagens, sombras, recuos). Típico: 0,65–0,75.",
        )
        st.session_state["packing_factor"] = pack

    if st.session_state.get("adv_map_center"):
        st.caption("3) Desenhe o polígono sobre o telhado/área disponível e confirme.")
        map_data = render_area_map(st.session_state["adv_map_center"])
        drawing = map_data.get("last_active_drawing") if isinstance(map_data, dict) else None
        if drawing and drawing.get("geometry", {}).get("type") == "Polygon":
            coords = drawing["geometry"]["coordinates"][0]
            area_m2 = calculate_polygon_area_m2([(c[1], c[0]) for c in coords[:-1]])
            st.session_state["available_area_m2"] = area_m2
            st.success(f"4) Área confirmada: **{area_m2:.2f} m²**")
            mark_step_complete(4)

    area_m2_manual = st.number_input(
        "Ou informe a área disponível manualmente (m²)",
        min_value=0.0,
        value=float(st.session_state.get("available_area_m2") or 0.0),
        step=1.0,
        help="Se não usar o mapa, preencha aqui a área útil do telhado ou solo.",
    )
    if area_m2_manual > 0:
        st.session_state["available_area_m2"] = area_m2_manual
        mark_step_complete(4)

    if st.session_state.get("available_area_m2"):
        area = float(st.session_state["available_area_m2"])
        approx_kwp = area * st.session_state.get("packing_factor", 0.7) / 2.3 * 0.55
        c1, c2 = st.columns(2)
        c1.metric("Área disponível (m²)", fmt_num(area, 2))
        c2.metric(
            "Potência máxima estimada (kWp)",
            fmt_num(approx_kwp, 2),
            help="Estimativa aproximada: depende do módulo selecionado na Etapa 5.",
        )

# ===========================================================================
# ETAPA 5 — Dimensionamento Técnico
# ===========================================================================
elif step == 5:
    st.header("ETAPA 5 — Dimensionamento Técnico")
    st.info("Nesta etapa, dimensionamos a melhor solução tecnicamente viável.")

    c1, c2 = st.columns(2)
    with c1:
        azimuth_manual = st.selectbox(
            "Azimute (orientação da face dos módulos)",
            [0, 90, 180, 270],
            index=2,
            format_func=lambda x: {0: "Norte (0°)", 90: "Leste (90°)", 180: "Sul (180°)", 270: "Oeste (270°)"}[x],
            help="No hemisfério Sul, 180° (norte geográfico) é o ideal para maximizar geração.",
        )
    with c2:
        tilt = st.number_input(
            "Inclinação dos módulos (°)",
            value=float(abs(st.session_state["latitude"])),
            format="%.2f",
            help="Ângulo de inclinação em relação ao plano horizontal. Padrão ≈ latitude local.",
        )

    if st.session_state.get("adv_map_center"):
        with st.expander("Calcular azimute pelo mapa (opcional)"):
            st.caption("Marque uma linha no mapa: ponto inicial na base dos módulos, ponto final na direção da face.")
            map_orient = render_area_map(st.session_state["adv_map_center"], key="azimuth_map")
            drawing2 = map_orient.get("last_active_drawing") if isinstance(map_orient, dict) else None
            if drawing2 and drawing2.get("geometry", {}).get("type") == "LineString":
                coords = drawing2["geometry"]["coordinates"]
                if len(coords) >= 2:
                    azimuth_map = calculate_azimuth_from_points(coords[0][1], coords[0][0], coords[-1][1], coords[-1][0])
                    st.session_state["azimuth_from_map"] = azimuth_map
                    st.success(f"Azimute calculado: {azimuth_map:.1f}°")

    azimuth = int(round(st.session_state.get("azimuth_from_map", azimuth_manual)))

    hist_kwh_year = 0.0
    if st.session_state["tariff_group"] == "B":
        hist_kwh_year = float(pd.to_numeric(st.session_state["historico_df"]["Consumo (kWh)"], errors="coerce").fillna(0).sum())
    elif "historico_df_a" in st.session_state:
        hda = st.session_state["historico_df_a"]
        hist_kwh_year = float(
            pd.to_numeric(hda["Energia ponta (kWh)"], errors="coerce").fillna(0).sum()
            + pd.to_numeric(hda["Energia fora ponta (kWh)"], errors="coerce").fillna(0).sum()
        )

    monthly_target = hist_kwh_year / 12 if hist_kwh_year > 0 else 5000
    st.info(f"**Alvo mensal de energia:** {fmt_num(monthly_target, 1)} kWh  |  **Alvo anual:** {fmt_num(hist_kwh_year, 0)} kWh")

    if st.button("▶ Executar dimensionamento técnico", type="primary"):
        with st.spinner("Consultando PVWatts e dimensionando sistema... aguarde."):
            resultados_df, erro = _dimensionamento_cached(
                monthly_target,
                float(st.session_state["latitude"]),
                float(st.session_state["longitude"]),
                int(azimuth),
                float(tilt),
            )
        if erro:
            st.error(erro)
        else:
            if st.session_state.get("available_area_m2"):
                resultados_df, area_meta = aplicar_restricao_area_resultados(
                    resultados_df,
                    available_area_m2=st.session_state["available_area_m2"],
                    packing_factor=st.session_state.get("packing_factor", 0.7),
                )
                st.session_state["area_meta"] = area_meta
            st.session_state["resultados_df_adv"] = resultados_df
            st.success("Dimensionamento concluído.")
            mark_step_complete(5)

    res = st.session_state.get("resultados_df_adv")
    if isinstance(res, pd.DataFrame) and not res.empty:
        best = res.iloc[0]
        if best.get("fonte_geracao") == "estimativa_fallback":
            pvwatts_detail = get_pvwatts_last_error()
            msg = "PVWatts indisponível. Dimensionamento calculado com estimativa offline de produtividade solar."
            if pvwatts_detail:
                msg += f" Detalhe: {pvwatts_detail}"
            st.warning(msg)

        area_meta = st.session_state.get("area_meta") or {"module_area_m2": 2.3, "warnings": []}
        area_req = int(best.get("sistema_num_total_paineis", 0)) * float(area_meta.get("module_area_m2", 2.3))
        potencia_kwp = float(best.get("sistema_potencia_total_w", 0)) / 1000.0
        energia_anual = float(best.get("energia_gerada_anual_kwh", 0) or 0)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Potência sugerida (kWp)", fmt_num(potencia_kwp, 2))
        c2.metric("Total de módulos", fmt_num(best.get("sistema_num_total_paineis", 0), 0))
        c3.metric("Área necessária (m²)", fmt_num(area_req, 2))
        c4.metric("Energia anual estimada (kWh)", fmt_num(energia_anual, 0))

        with st.expander("Configuração técnica detalhada"):
            col_map = {
                "inversor_modelo": "Inversor",
                "inversor_fabricante": "Fabricante Inversor",
                "inversor_num_unidades": "Qtd. Inversores",
                "inversor_num_mppt": "MPPTs",
                "painel_modelo": "Painel",
                "painel_fabricante": "Fabricante Painel",
                "painel_potencia": "Potência Painel (Wp)",
                "arranjo_modulos_serie": "Módulos em série",
                "arranjo_conjuntos_paralelo_por_mppt": "Strings paralelas/MPPT",
                "sistema_num_total_paineis": "Total de painéis",
                "sistema_potencia_total_w": "Potência total (W)",
                "energia_gerada_anual_kwh": "Energia anual (kWh)",
                "fonte_geracao": "Fonte do cálculo",
            }
            detail_data = {col_map.get(k, k): best.get(k) for k in col_map if k in best}
            st.table(pd.DataFrame(detail_data.items(), columns=["Parâmetro", "Valor"]))
            for w in area_meta.get("warnings", []):
                st.warning(w)

        if len(res) > 1:
            with st.expander("Comparativo dos top sistemas encontrados"):
                st.dataframe(res.head(6), use_container_width=True)

# ===========================================================================
# ETAPA 6 — Análise Energética Avançada
# ===========================================================================
elif step == 6:
    st.header("ETAPA 6 — Análise Energética Avançada")
    st.info("Avaliamos como a geração FV interage com a operação real do cliente hora a hora.")

    res = st.session_state.get("resultados_df_adv")
    fit_adv = st.session_state.get("fit_adv")
    if res is None or fit_adv is None:
        st.warning("Execute o fit de perfil (Etapa 3) e o dimensionamento (Etapa 5) antes.")
    else:
        best = res.iloc[0]
        potencia_kwp = float(best["sistema_potencia_total_w"]) / 1000.0
        lat = float(st.session_state["latitude"])

        with st.spinner("Obtendo geração horária do PVWatts... aguarde."):
            pv_hourly = _geracao_horaria_cached(
                potencia_kwp, lat,
                float(st.session_state["longitude"]),
                180, float(abs(lat)),
            )

        load_day = fit_adv["load_curve_kw"]
        pv_day = pv_hourly[:24]

        peak_labels = classify_peak_hours(list(range(24)), "18:00", "21:00", weekdays_only=False)
        sc = analyze_self_consumption(load_day, pv_day)
        ps = analyze_peak_shaving(load_day, pv_day, peak_labels)
        ls = simulate_simplified_load_shifting(
            energy_peak_kwh=fit_adv.get("energy_peak_kwh_estimated", 0.0),
            energy_offpeak_kwh=fit_adv.get("energy_offpeak_kwh_estimated", 0.0),
            percentual_carga_deslocavel_ponta=0.15,
            shift_window="horário solar",
            tarifa_ponta=1.2,
            tarifa_fora_ponta=0.7,
        )
        st.session_state["advanced_metrics"] = {"self": sc, "peak": ps, "shift": ls}
        mark_step_complete(6)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Autoconsumo (%)", fmt_num(sc["self_consumption_ratio"] * 100, 1), help="% da geração FV consumida localmente.")
        c2.metric("Autossuficiência (%)", fmt_num(sc["self_sufficiency_ratio"] * 100, 1), help="% da carga atendida pela geração FV.")
        c3.metric("Injeção na rede (kWh/dia)", fmt_num(sc["total_grid_export_kwh"], 2))
        c4.metric("Peak shaving (kW)", fmt_num(ps["peak_shaving_kw"], 2), help="Redução de pico de demanda atingida pela geração FV.")

        horas = list(range(24))
        fig_energy = go.Figure()
        fig_energy.add_scatter(x=horas, y=load_day, name="Carga total (kW)", fill="tozeroy", line={"color": "#e74c3c"}, fillcolor="rgba(231,76,60,0.15)")
        fig_energy.add_scatter(x=horas, y=pv_day, name="Geração FV (kW)", fill="tozeroy", line={"color": "#f1c40f"}, fillcolor="rgba(241,196,15,0.2)")
        fig_energy.add_scatter(x=horas, y=ps["net_load_kw"], name="Carga líquida (kW)", line={"color": "#2ecc71", "dash": "dash"})
        fig_energy.update_layout(
            title="Balanço energético diário típico",
            xaxis_title="Hora",
            yaxis_title="kW",
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font_color="#eef4ff",
            legend={"orientation": "h"},
        )
        st.plotly_chart(fig_energy, use_container_width=True)

        with st.expander("Geração horária anual (8760h)"):
            pv_annual_df = pd.DataFrame({"Hora do ano": list(range(len(pv_hourly))), "Geração FV (kW)": pv_hourly})
            fig_ann = px.area(pv_annual_df, x="Hora do ano", y="Geração FV (kW)", title="Perfil anual de geração FV", color_discrete_sequence=["#f1c40f"])
            fig_ann.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", font_color="#eef4ff")
            st.plotly_chart(fig_ann, use_container_width=True)

        with st.expander("Painel técnico detalhado"):
            st.json({"autoconsumo": sc, "peak_shaving": ps, "load_shifting": ls})

# ===========================================================================
# ETAPA 7 — Viabilidade Econômica
# ===========================================================================
elif step == 7:
    st.header("ETAPA 7 — Viabilidade Econômica")
    st.info("Esta etapa traduz engenharia em decisão financeira com projeção de 25 anos.")

    adv = st.session_state.get("advanced_metrics")
    res = st.session_state.get("resultados_df_adv")
    if adv is None or res is None:
        st.warning("Execute as etapas 5 e 6 antes.")
    else:
        best = res.iloc[0]
        annual_generation = float(best.get("energia_gerada_anual_kwh", 0.0) or 0.0)

        st.markdown("#### Parâmetros de investimento")
        capex_modo = st.radio(
            "Modo de entrada do CAPEX",
            ["Simplificado (valor único)", "Detalhado (por categoria)"],
            horizontal=True,
            help="Simplificado: informe só o valor total. Detalhado: distribua por categoria para o relatório.",
        )
        c1, c2 = st.columns(2)
        with c1:
            if capex_modo == "Simplificado (valor único)":
                st.markdown("**CAPEX total do projeto (R$)**")
                capex_total = st.number_input(
                    "Valor total do investimento (R$)",
                    min_value=0.0, value=300000.0, step=1000.0, format="%.2f",
                    help="Informe o valor final do projeto incluindo equipamentos, instalação e engenharia.",
                )
                capex_paineis = capex_total * 0.50
                capex_inversores = capex_total * 0.20
                capex_estrutura = capex_total * 0.10
                capex_eletrica = capex_total * 0.08
                capex_mao_obra = capex_total * 0.08
                capex_engenharia = capex_total * 0.04
                st.caption("Distribuição estimada gerada automaticamente para o relatório.")
            else:
                st.markdown("**CAPEX detalhado por categoria (R$)**")
                capex_paineis = st.number_input("Painéis solares", min_value=0.0, value=150000.0, step=1000.0, format="%.2f")
                capex_inversores = st.number_input("Inversores", min_value=0.0, value=60000.0, step=1000.0, format="%.2f")
                capex_estrutura = st.number_input("Estrutura e fixação", min_value=0.0, value=30000.0, step=1000.0, format="%.2f")
                capex_eletrica = st.number_input("Elétrica e cabeamento", min_value=0.0, value=25000.0, step=1000.0, format="%.2f")
                capex_mao_obra = st.number_input("Mão de obra / instalação", min_value=0.0, value=25000.0, step=1000.0, format="%.2f")
                capex_engenharia = st.number_input("Engenharia / documentação", min_value=0.0, value=10000.0, step=1000.0, format="%.2f")
                capex_total = capex_paineis + capex_inversores + capex_estrutura + capex_eletrica + capex_mao_obra + capex_engenharia
            st.metric("**CAPEX total (R$)**", fmt_num(capex_total, 2))

        with c2:
            st.markdown("**Premissas econômicas**")
            tariff = st.number_input(
                "Tarifa base de energia (R$/kWh)",
                min_value=0.1, value=0.85, step=0.01,
                help="Tarifa média atual do cliente, incluindo tributos.",
            )
            opex_anual = st.number_input(
                "OPEX anual estimado (R$)",
                min_value=0.0, value=2000.0, step=500.0,
                help="Custo de operação e manutenção anual (limpeza, seguros, monitoramento).",
            )
            degradacao = st.slider(
                "Degradação anual dos painéis (%)",
                min_value=0.1, max_value=1.5, value=0.6, step=0.1,
                help="Perda anual de eficiência dos painéis. Fabricantes garantem tipicamente 0,5–0,7%/ano.",
            ) / 100
            escalada_tarifa = st.slider(
                "Escalada tarifária anual (%)",
                min_value=0.0, max_value=15.0, value=5.0, step=0.5,
                help="Crescimento anual esperado da tarifa elétrica. Histórico ANEEL: 5–8%/ano.",
            ) / 100
            taxa_desconto = st.slider(
                "Taxa de desconto / WACC (%)",
                min_value=1.0, max_value=20.0, value=10.0, step=0.5,
                help="Custo de capital para cálculo do VPL. Use a TMA da empresa ou SELIC.",
            ) / 100

        # Cálculo do fluxo de caixa
        fluxo = calcular_fluxo_caixa(
            capex=capex_total,
            geracao_ano1_kwh=annual_generation,
            tarifa_base=tariff,
            degradacao_anual=degradacao,
            escalada_tarifa=escalada_tarifa,
            anos=25,
            taxa_desconto=taxa_desconto,
            opex_anual_rs=opex_anual,
        )

        annual_savings_yr1 = annual_generation * tariff
        payback_simples = capex_total / annual_savings_yr1 if annual_savings_yr1 > 0 else None
        payback_desc = calcular_payback_descontado(fluxo)
        vpn_val = calcular_vpn(capex_total, fluxo)
        tir_real = calcular_tir_real(
            capex_total, annual_generation, tariff,
            degradacao_anual=degradacao, escalada_tarifa=escalada_tarifa,
            anos=25, opex_anual_rs=opex_anual,
        )
        geracao_25 = sum(r["Geração (kWh)"] for r in fluxo)
        economia_total_25 = sum(r["Fluxo líquido (R$)"] for r in fluxo)
        roi_25 = (economia_total_25 - capex_total) / capex_total * 100 if capex_total > 0 else None
        co2 = calcular_co2_evitado(geracao_25)

        eco_metrics = {
            "economia_anual": annual_savings_yr1,
            "capex_total": capex_total,
            "payback": payback_simples,
            "payback_descontado": payback_desc,
            "vpn": vpn_val,
            "tir_real": tir_real,
            "roi": roi_25,
            "fluxo_caixa_25anos": fluxo,
            "co2": co2,
            "sensibilidade_tarifaria": {
                "-20%": annual_generation * tariff * 0.8,
                "-10%": annual_generation * tariff * 0.9,
                "base": annual_savings_yr1,
                "+10%": annual_generation * tariff * 1.1,
                "+20%": annual_generation * tariff * 1.2,
            },
            "capex_breakdown": {
                "Painéis": capex_paineis,
                "Inversores": capex_inversores,
                "Estrutura": capex_estrutura,
                "Elétrica": capex_eletrica,
                "Mão de obra": capex_mao_obra,
                "Engenharia": capex_engenharia,
            },
        }
        st.session_state["economic_metrics"] = eco_metrics
        mark_step_complete(7)

        st.markdown("#### Indicadores financeiros")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Economia ano 1 (R$)", fmt_num(annual_savings_yr1, 2))
        c2.metric("Payback simples (anos)", fmt_num(payback_simples, 1))
        c3.metric("Payback descontado (anos)", fmt_num(payback_desc, 1) if payback_desc else "Não atingido")
        c4.metric("VPL 25 anos (R$)", fmt_num(vpn_val, 2))
        c5.metric("TIR real (%)", fmt_num((tir_real or 0) * 100, 2))

        c6, c7 = st.columns(2)
        c6.metric("ROI 25 anos (%)", fmt_num(roi_25, 1))
        c7.metric("CO₂ evitado 25 anos (t)", fmt_num(co2["co2_evitado_t"], 1))

        st.markdown(
            f'<div class="co2-badge">🌱 Equivalente a <strong>{fmt_num(co2["arvores_equivalentes"], 0)} árvores</strong> plantadas — '
            f'<strong>{fmt_num(co2["co2_evitado_kg"], 0)} kg</strong> de CO₂ evitados ao longo de 25 anos.</div>',
            unsafe_allow_html=True,
        )

        st.markdown("#### Gráficos financeiros")
        tab1, tab2, tab3, tab4 = st.tabs(["VPL Acumulado", "Fluxo de Caixa", "Sensibilidade Tarifária", "CAPEX Breakdown"])

        with tab1:
            df_fc = pd.DataFrame(fluxo)
            fig_vpn = go.Figure()
            fig_vpn.add_scatter(
                x=df_fc["Ano"], y=df_fc["VPL acumulado (R$)"],
                name="VPL acumulado", fill="tozeroy",
                line={"color": "#2ecc71"},
                fillcolor="rgba(46,204,113,0.15)",
            )
            fig_vpn.add_hline(y=0, line_dash="dash", line_color="#e74c3c", annotation_text="Break-even")
            if payback_desc:
                fig_vpn.add_vline(x=payback_desc, line_dash="dot", line_color="#f1c40f",
                                  annotation_text=f"Payback descontado: {payback_desc:.1f} anos")
            fig_vpn.update_layout(
                title="Evolução do VPL acumulado (25 anos)",
                xaxis_title="Ano",
                yaxis_title="R$",
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
                font_color="#eef4ff",
            )
            st.plotly_chart(fig_vpn, use_container_width=True)

        with tab2:
            fig_fc = go.Figure()
            fig_fc.add_bar(x=df_fc["Ano"], y=df_fc["Fluxo líquido (R$)"], name="Fluxo bruto", marker_color="#2c7be5")
            fig_fc.add_scatter(x=df_fc["Ano"], y=df_fc["Geração (kWh)"], name="Geração (kWh)", yaxis="y2", line={"color": "#f1c40f"})
            fig_fc.update_layout(
                title="Fluxo de caixa anual + geração ao longo dos 25 anos",
                xaxis_title="Ano",
                yaxis_title="R$",
                yaxis2={"title": "kWh", "overlaying": "y", "side": "right"},
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
                font_color="#eef4ff",
                legend={"orientation": "h"},
            )
            st.plotly_chart(fig_fc, use_container_width=True)

        with tab3:
            sens = eco_metrics["sensibilidade_tarifaria"]
            fig_sens = px.bar(
                x=list(sens.keys()), y=list(sens.values()),
                labels={"x": "Cenário tarifário", "y": "Economia ano 1 (R$)"},
                title="Sensibilidade da economia ao preço da energia",
                color=list(sens.values()),
                color_continuous_scale="RdYlGn",
            )
            fig_sens.update_layout(
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
                font_color="#eef4ff",
                coloraxis_showscale=False,
            )
            st.plotly_chart(fig_sens, use_container_width=True)

        with tab4:
            breakdown = eco_metrics["capex_breakdown"]
            fig_pie = px.pie(
                names=list(breakdown.keys()),
                values=list(breakdown.values()),
                title="Composição do CAPEX",
                color_discrete_sequence=px.colors.sequential.Blues_r,
            )
            fig_pie.update_layout(
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
                font_color="#eef4ff",
            )
            st.plotly_chart(fig_pie, use_container_width=True)

        with st.expander("Tabela de fluxo de caixa ano a ano"):
            st.dataframe(
                df_fc.style.format({
                    "Geração (kWh)": "{:,.0f}",
                    "Tarifa (R$/kWh)": "{:.4f}",
                    "Economia bruta (R$)": "R$ {:,.2f}",
                    "Fluxo líquido (R$)": "R$ {:,.2f}",
                    "Fluxo descontado (R$)": "R$ {:,.2f}",
                    "VPL acumulado (R$)": "R$ {:,.2f}",
                }),
                use_container_width=True,
            )

# ===========================================================================
# ETAPA 8 — Relatório Final
# ===========================================================================
elif step == 8:
    st.header("ETAPA 8 — Relatório Final")
    st.info("Consolidação técnica + executiva em formato premium para apresentação ao cliente.")

    res = st.session_state.get("resultados_df_adv")
    eco = st.session_state.get("economic_metrics", {}) or {}
    adv = st.session_state.get("advanced_metrics", {}) or {}

    report = {
        "1_resumo_executivo": {
            "cliente": st.session_state.get("cliente_nome", "N/D"),
            "contexto": f"{st.session_state.get('cliente_tipo', 'N/D')} | Grupo {st.session_state.get('tariff_group', 'N/D')}",
            "local": st.session_state.get("localizacao", "N/D"),
            "distribuidora": st.session_state.get("distribuidora", "N/D"),
        },
        "2_premissas": {
            "latitude": st.session_state.get("latitude"),
            "longitude": st.session_state.get("longitude"),
            "modalidade": st.session_state.get("modalidade"),
        },
        "3_historico_consumo": (
            st.session_state.get("historico_df_a", st.session_state.get("historico_df")).to_dict(orient="records")
            if isinstance(st.session_state.get("historico_df_a", st.session_state.get("historico_df")), pd.DataFrame)
            else []
        ),
        "4_curva_demanda": st.session_state.get("fit_adv", {}),
        "5_area_disponivel": {
            "area_m2": st.session_state.get("available_area_m2"),
            "fator_aproveitamento": st.session_state.get("packing_factor", 0.7),
        },
        "6_sistema_proposto": (
            res.iloc[0].to_dict() if isinstance(res, pd.DataFrame) and not res.empty else {}
        ),
        "7_analise_energetica": adv,
        "8_viabilidade_economica": eco,
        "9_limitacoes": (st.session_state.get("area_meta") or {}).get("warnings", []),
        "10_proximos_passos": [
            "Validar curva com medição real de carga (15 min).",
            "Refinar premissas tarifárias da distribuidora local.",
            "Executar visita técnica para engenharia de instalação.",
        ],
    }

    # --- Cards de resumo executivo ---
    st.subheader("Resumo do Projeto")
    ex = report["1_resumo_executivo"]
    prem = report["2_premissas"]
    c1, c2, c3, c4 = st.columns(4)
    c1.info(f"**Cliente:** {ex.get('cliente', 'N/D')}")
    c2.info(f"**Local:** {ex.get('local', 'N/D')}")
    c3.info(f"**Contexto:** {ex.get('contexto', 'N/D')}")
    c4.info(f"**Distribuidora:** {ex.get('distribuidora', 'N/D')}")

    # --- KPIs principais ---
    if eco:
        st.markdown("#### Indicadores-chave")
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        sistema = report["6_sistema_proposto"]
        c1.metric("Potência (kWp)", fmt_num(float(sistema.get("sistema_potencia_total_w", 0)) / 1000, 2))
        c2.metric("Geração anual (kWh)", fmt_num(float(sistema.get("energia_gerada_anual_kwh", 0) or 0), 0))
        c3.metric("Economia ano 1 (R$)", fmt_num(eco.get("economia_anual", 0), 2))
        c4.metric("Payback (anos)", fmt_num(eco.get("payback", 0), 1))
        c5.metric("VPL 25 anos (R$)", fmt_num(eco.get("vpn", 0), 2))
        c6.metric("TIR real (%)", fmt_num((eco.get("tir_real") or 0) * 100, 2))

        co2 = eco.get("co2", {}) or {}
        if co2:
            st.markdown(
                f'<div class="co2-badge">🌱 Impacto ambiental 25 anos: <strong>{fmt_num(co2.get("co2_evitado_t", 0), 1)} tCO₂</strong> evitados '
                f'— equivalente a <strong>{fmt_num(co2.get("arvores_equivalentes", 0), 0)} árvores</strong>.</div>',
                unsafe_allow_html=True,
            )

    # --- Gráfico VPL no relatório ---
    fluxo = eco.get("fluxo_caixa_25anos", [])
    if fluxo:
        df_fc = pd.DataFrame(fluxo)
        fig_rep = go.Figure()
        fig_rep.add_scatter(
            x=df_fc["Ano"], y=df_fc["VPL acumulado (R$)"],
            name="VPL", fill="tozeroy",
            line={"color": "#2ecc71"},
            fillcolor="rgba(46,204,113,0.15)",
        )
        fig_rep.add_hline(y=0, line_dash="dash", line_color="#e74c3c")
        fig_rep.update_layout(
            title="VPL acumulado ao longo dos 25 anos",
            xaxis_title="Ano",
            yaxis_title="R$",
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font_color="#eef4ff",
        )
        st.plotly_chart(fig_rep, use_container_width=True)

    # --- Memória de cálculo ---
    with st.expander("📋 Memória de cálculo (narrativa técnica)", expanded=False):
        lat = st.session_state.get("latitude")
        lon = st.session_state.get("longitude")
        limites = (st.session_state.get("area_meta") or {}).get("warnings", [])
        st.markdown(f"""
**Narrativa técnica**

1. Com base nos {len(report['3_historico_consumo'])} meses informados, estimou-se o consumo anual do cliente.
2. A curva típica selecionada foi calibrada para representar o comportamento energético mais provável.
3. A área disponível ({fmt_num(st.session_state.get('available_area_m2'), 2)} m²) limita a potência máxima instalável.
4. O sistema foi dimensionado via PVWatts com dados reais de irradiação solar e equipamentos cadastrados.
5. O modelo financeiro considera degradação anual dos painéis e escalada tarifária composta ao longo de 25 anos.

**Alertas e premissas**
- Localização: lat {fmt_num(lat, 4)} / lon {fmt_num(lon, 4)}
- Grupo tarifário: {st.session_state.get('tariff_group', 'N/D')}
- Limitações: {', '.join(limites) if limites else 'Sem alertas críticos'}
""")

    # --- Downloads ---
    st.markdown("#### Exportar relatório")
    dcol1, dcol2, dcol3 = st.columns(3)

    with dcol1:
        report_json = json.dumps(report, ensure_ascii=False, indent=2, default=str)
        st.download_button(
            "📄 Baixar JSON completo",
            report_json,
            file_name="pace_relatorio_final.json",
            mime="application/json",
            use_container_width=True,
        )

    with dcol2:
        zip_bytes = build_latex_report_zip(report)
        st.download_button(
            "📦 Baixar pacote LaTeX (.zip)",
            zip_bytes,
            file_name="pace_relatorio_latex.zip",
            mime="application/zip",
            use_container_width=True,
        )

    with dcol3:
        _hdf = st.session_state.get("historico_df_a")
        if not isinstance(_hdf, pd.DataFrame) or _hdf.empty:
            _hdf = st.session_state.get("historico_df")
        historico_df = _hdf if isinstance(_hdf, pd.DataFrame) else pd.DataFrame()
        sistema_dict = report["6_sistema_proposto"]
        premissas_dict = {
            "Cliente": st.session_state.get("cliente_nome", ""),
            "Local": st.session_state.get("localizacao", ""),
            "Latitude": lat,
            "Longitude": lon,
            "Grupo tarifário": st.session_state.get("tariff_group", ""),
            "Modalidade": st.session_state.get("modalidade", ""),
            "Área disponível (m²)": st.session_state.get("available_area_m2", ""),
            "Fator aproveitamento": st.session_state.get("packing_factor", 0.7),
        }
        if eco and fluxo:
            excel_bytes = gerar_excel_completo(
                historico_df=historico_df,
                resultados_dim=sistema_dict,
                fluxo_caixa=fluxo,
                eco_metrics=eco,
                co2=eco.get("co2", {}),
                premissas=premissas_dict,
            )
            st.download_button(
                "📊 Baixar Excel completo (.xlsx)",
                excel_bytes,
                file_name="pace_relatorio_completo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.info("Complete as etapas 5–7 para gerar o Excel.")

    mark_step_complete(8)

# ---------------------------------------------------------------------------
# Navegação inferior
# ---------------------------------------------------------------------------
st.markdown("---")
nav_prev, nav_space, nav_next = st.columns([1, 4, 1])
with nav_prev:
    if st.button("⬅️ Voltar", disabled=st.session_state["current_step"] <= 1, use_container_width=True):
        st.session_state["current_step"] -= 1
        st.rerun()
with nav_next:
    if st.button("Próximo ➡️", disabled=st.session_state["current_step"] >= 8, use_container_width=True, type="primary"):
        mark_step_complete(st.session_state["current_step"])
        st.session_state["current_step"] += 1
        st.rerun()
