import io
import json
import os
import zipfile
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

from pv_calculator import (
    aplicar_restricao_area_resultados,
    calcular_geracao_horaria_pvwatts,
    carregar_dados_equipamentos,
    geocode_location,
    realizar_dimensionamento_completo,
    salvar_novo_inversor,
    salvar_novo_painel,
)
from src.client_db import list_projects, load_project, upsert_project
from src.load_profile_fitting import fit_load_profile_to_bill
from src.load_profiles import load_typical_profiles
from src.map_area import calculate_polygon_area_m2, calculate_azimuth_from_points, geocode_address, render_area_map
from src.peak_shaving import analyze_peak_shaving
from src.self_consumption import analyze_self_consumption
from src.tariff_periods import classify_peak_hours
from src.load_shifting import simulate_simplified_load_shifting


def build_latex_report_zip(report: dict) -> bytes:
    files = {}
    figures_tex = []
    historico = report.get("3_historico_consumo", [])
    if historico:
        meses = [str(row.get("Mês", i + 1)) for i, row in enumerate(historico)]
        consumo = [float(row.get("Consumo (kWh)", 0) or 0) for row in historico]
        coords = " ".join([f"({m},{v:.2f})" for m, v in zip(meses, consumo)])
        figures_tex.append(
            "\\begin{figure}[h!]\\centering\\begin{tikzpicture}\\begin{axis}[ybar,symbolic x coords={"
            + ",".join(meses)
            + "},xtick=data,x tick label style={rotate=45,anchor=east},width=0.95\\linewidth,height=6cm,ylabel={kWh},title={Histórico de consumo}]"
            + f"\\addplot coordinates {{{coords}}};"
            + "\\end{axis}\\end{tikzpicture}\\caption{Histórico de consumo.}\\end{figure}"
        )

    eco = report.get("8_viabilidade_economica", {}) or {}
    sens = eco.get("sensibilidade_tarifaria", {})
    if sens:
        labels = list(sens.keys())
        values = [float(v or 0) for v in sens.values()]
        coords = " ".join([f"({m},{v:.2f})" for m, v in zip(labels, values)])
        figures_tex.append(
            "\\begin{figure}[h!]\\centering\\begin{tikzpicture}\\begin{axis}[ybar,symbolic x coords={"
            + ",".join(labels)
            + "},xtick=data,width=0.75\\linewidth,height=6cm,ylabel={R\\$},title={Sensibilidade tarifária}]"
            + f"\\addplot coordinates {{{coords}}};"
            + "\\end{axis}\\end{tikzpicture}\\caption{Sensibilidade tarifária.}\\end{figure}"
        )

    if not figures_tex:
        figures_tex.append("Sem dados suficientes para gerar gráficos.")

    latex = f"""\\documentclass[11pt,a4paper]{{article}}
\\usepackage[utf8]{{inputenc}}
\\usepackage[T1]{{fontenc}}
\\usepackage[brazil]{{babel}}
\\usepackage{{graphicx}}
\\usepackage{{booktabs}}
\\usepackage{{geometry}}
\\usepackage{{pgfplots}}
\\pgfplotsset{{compat=1.18}}
\\geometry{{margin=2.2cm}}
\\title{{Relatório Técnico Fotovoltaico - AurumCalc}}
\\author{{AurumCalc}}
\\date{{\\today}}
\\begin{{document}}
\\maketitle
\\section{{Resumo executivo}}
Cliente: {report.get("1_resumo_executivo", {}).get("cliente", "N/D")}\\\\
Local: {report.get("1_resumo_executivo", {}).get("local", "N/D")}\\\\
\\section{{Premissas}}
Latitude: {report.get("2_premissas", {}).get("latitude", "N/D")}\\\\
Longitude: {report.get("2_premissas", {}).get("longitude", "N/D")}\\\\
\\section{{Análises gráficas}}
{"".join(figures_tex)}
\\section{{Resultados principais}}
Economia anual (R$): {eco.get("economia_anual", "N/D")}\\\\
Payback (anos): {eco.get("payback", "N/D")}\\\\
ROI 25 anos (\\%): {eco.get("roi", "N/D")}\\\\
\\section{{Próximos passos}}
\\begin{{itemize}}
\\item Validar curva com medição real.
\\item Refinar premissas tarifárias.
\\item Realizar visita técnica.
\\end{{itemize}}
\\end{{document}}
"""
    files["relatorio.tex"] = latex.encode("utf-8")
    files["relatorio_dados.json"] = json.dumps(report, ensure_ascii=False, indent=2, default=str).encode("utf-8")

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, data in files.items():
            zf.writestr(fname, data)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

if "PVWATTS_API_KEY" in st.secrets:
    os.environ["PVWATTS_API_KEY"] = st.secrets["PVWATTS_API_KEY"]

st.set_page_config(page_title="AurumCalc | Diagnóstico Energético", layout="wide", initial_sidebar_state="expanded")

st.markdown(
    """
<style>
.stApp { background: linear-gradient(180deg, #061127 0%, #0b1f3a 100%); color: #eef4ff; }
section[data-testid="stSidebar"] { background: #081630; }
.card { background: #0f2546; border-radius: 14px; padding: 1rem; border: 1px solid #264d86; }
.kpi { background: #132b52; border-radius: 12px; padding: 0.7rem; border: 1px solid #2c5fa5; }
.step-pill { background: #123160; color: #d9e8ff; padding: 0.35rem 0.7rem; border-radius: 999px; font-size: 0.82rem; }
hr { border-color: #1f3f70; }
</style>
""",
    unsafe_allow_html=True,
)

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


def fmt_num(x, casas=2, default="-"):
    try:
        s = f"{float(x):,.{casas}f}"
        return s.replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return default


def irr_bisection(initial_investment: float, annual_cashflow: float, years: int = 20):
    if initial_investment <= 0 or annual_cashflow <= 0:
        return None
    lo, hi = -0.99, 1.5
    for _ in range(120):
        mid = (lo + hi) / 2
        npv = -initial_investment + sum(annual_cashflow / ((1 + mid) ** t) for t in range(1, years + 1))
        if npv > 0:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


def load_data():
    try:
        return carregar_dados_equipamentos(FILE_PATH_EQUIPAMENTOS)
    except Exception:
        return pd.DataFrame(), pd.DataFrame()


def ensure_state():
    defaults = {
        "current_step": 1,
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
        else:
            payload[key] = value
    return payload


def _restore_state(payload):
    for key, value in payload.items():
        if isinstance(value, dict) and value.get("__type__") == "dataframe":
            st.session_state[key] = pd.DataFrame(value.get("records", []))
        else:
            st.session_state[key] = value

# Sidebar
st.sidebar.title("AurumCalc • Jornada Guiada")
progress = (st.session_state["current_step"] - 1) / (len(STEP_NAMES) - 1)
st.sidebar.progress(progress, text=f"Progresso: etapa {st.session_state['current_step']} de {len(STEP_NAMES)}")
for i, step in enumerate(STEP_NAMES, start=1):
    icon = "✅" if i < st.session_state["current_step"] else ("🟦" if i == st.session_state["current_step"] else "⬜")
    st.sidebar.markdown(f"{icon} {step}")

with st.sidebar.expander("Base de equipamentos"):
    df_p, df_i = load_data()
    st.caption(f"Painéis cadastrados: {len(df_p)}")
    st.caption(f"Inversores cadastrados: {len(df_i)}")

with st.sidebar.expander("💾 Banco de projetos por cliente", expanded=True):
    db_client = st.text_input("Cliente para buscar/salvar", value=st.session_state.get("cliente_nome", ""))
    if db_client and not st.session_state.get("cliente_nome"):
        st.session_state["cliente_nome"] = db_client

    project_rows = list_projects(db_client if db_client else None)
    project_labels = [f"{row[0]} | {row[1]} | {row[2][:19]}" for row in project_rows]
    selected_label = st.selectbox("Projetos salvos", options=["Selecione..."] + project_labels)

    if st.button("Carregar projeto selecionado"):
        if selected_label == "Selecione...":
            st.warning("Selecione um projeto salvo para carregar.")
        else:
            idx = project_labels.index(selected_label)
            client_name, project_name, _ = project_rows[idx]
            payload = load_project(client_name, project_name)
            if payload:
                _restore_state(payload)
                st.session_state["active_project"] = f"{client_name}:{project_name}"
                st.success(f"Projeto '{project_name}' carregado para o cliente '{client_name}'.")
                st.rerun()
            else:
                st.error("Projeto não encontrado no banco de dados.")

    st.session_state["project_name"] = st.text_input(
        "Nome do projeto",
        value=st.session_state.get("project_name", "") or st.session_state.get("localizacao", ""),
    )
    if st.button("Salvar/Atualizar projeto no banco"):
        client_name = (db_client or st.session_state.get("cliente_nome", "")).strip()
        project_name = st.session_state.get("project_name", "").strip()
        if not client_name or not project_name:
            st.error("Informe cliente e nome do projeto para salvar.")
        else:
            upsert_project(client_name, project_name, _serialize_state())
            st.session_state["active_project"] = f"{client_name}:{project_name}"
            st.success(f"Projeto '{project_name}' salvo com sucesso para o cliente '{client_name}'.")

with st.sidebar.expander("➕ Novo equipamento"):
    tab_p, tab_i = st.tabs(["Painel", "Inversor"])
    with tab_p:
        with st.form("novo_painel", clear_on_submit=True):
            modelo_p = st.text_input("Modelo")
            fabricante_p = st.text_input("Fabricante")
            pmax = st.number_input("Pmax (Wp)", min_value=1.0, value=550.0)
            voc = st.number_input("Voc (V)", min_value=1.0, value=49.5)
            isc = st.number_input("Isc (A)", min_value=0.1, value=13.4)
            submit_p = st.form_submit_button("Salvar painel")
            if submit_p:
                ok = salvar_novo_painel(
                    {
                        "modelo": modelo_p,
                        "fabricante": fabricante_p,
                        "potencia_maxima_nominal_pmax": pmax,
                        "tensao_operacao_otima_vmp": 0,
                        "corrente_operacao_otima_imp": 0,
                        "tensao_circuito_aberto_voc": voc,
                        "corrente_curto_circuito_isc": isc,
                        "eficiencia_modulo": 0,
                    }
                )
                st.success("Painel salvo.") if ok else st.error("Falha ao salvar painel.")
    with tab_i:
        with st.form("novo_inversor", clear_on_submit=True):
            modelo_i = st.text_input("Modelo ")
            fabricante_i = st.text_input("Fabricante ")
            pot_ca = st.number_input("Potência CA (W)", min_value=1.0, value=50000.0)
            vmax = st.number_input("Vmax CC (V)", min_value=1.0, value=1100.0)
            submit_i = st.form_submit_button("Salvar inversor")
            if submit_i:
                ok = salvar_novo_inversor(
                    {
                        "modelo": modelo_i,
                        "fabricante": fabricante_i,
                        "potencia_maxima_fv_maxima": pot_ca,
                        "tensao_maxima_cc": vmax,
                        "tensao_start": 200,
                        "tensao_nominal": 600,
                        "faixa_tensao_mpp": "200-850",
                        "numero_mpp_trackers": 4,
                        "corrente_maxima_entrada_por_mppt_tracker": 30,
                        "corrente_maxima_curto_circuito_por_mppt_tracker": 40,
                        "maxima_potencia_nominal_ca": pot_ca,
                        "potencia_maxima_aparente_ca": pot_ca,
                        "tensao_nominal_ca": 380,
                        "frequencia_rede_ca": "60Hz",
                        "corrente_saida_maxima": 80,
                        "fator_potencia_ajustavel": "0.8i-0.8c",
                        "quantidade_fases_ca": 3,
                    }
                )
                st.success("Inversor salvo.") if ok else st.error("Falha ao salvar inversor.")

st.title("⚡ AurumCalc — Plataforma Premium de Diagnóstico Energético")
st.caption("Do diagnóstico ao relatório final: uma jornada consultiva, guiada e estratégica.")

step = st.session_state["current_step"]
st.markdown(f"<span class='step-pill'>Etapa atual: {STEP_NAMES[step-1]}</span>", unsafe_allow_html=True)
st.markdown("---")

if step == 1:
    st.header("ETAPA 1 — Identificação do Projeto")
    st.info("Vamos primeiro entender quem é o consumidor e qual é o contexto energético do projeto.")
    c1, c2 = st.columns(2)
    with c1:
        st.session_state["cliente_nome"] = st.text_input("Nome do cliente", value=st.session_state.get("cliente_nome", ""))
        st.session_state["cliente_tipo"] = st.selectbox("Tipo de cliente", ["Comercial", "Industrial", "Rural", "Residencial", "Poder Público"])
        st.session_state["tariff_group"] = st.radio("Grupo tarifário", ["A", "B"], horizontal=True)
        st.session_state["modalidade"] = st.selectbox("Modalidade tarifária", ["Convencional", "Branca", "Verde", "Azul"])
    with c2:
        st.session_state["localizacao"] = st.text_input("Localização", value=st.session_state.get("localizacao", "São Paulo, SP"))
        if st.button("Buscar coordenadas por localização"):
            lat_geo, lon_geo = geocode_location(st.session_state["localizacao"])
            if lat_geo is not None:
                st.session_state["latitude"], st.session_state["longitude"] = float(lat_geo), float(lon_geo)
                st.success("Coordenadas aplicadas.")
            else:
                st.error("Não foi possível localizar o endereço.")
        st.session_state["latitude"] = st.number_input("Latitude", value=float(st.session_state["latitude"]), format="%.6f")
        st.session_state["longitude"] = st.number_input("Longitude", value=float(st.session_state["longitude"]), format="%.6f")
        st.session_state["distribuidora"] = st.text_input("Distribuidora", value=st.session_state.get("distribuidora", ""))
    st.session_state["observacoes"] = st.text_area("Observações do projeto", value=st.session_state.get("observacoes", ""))

if step == 2:
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
        est_count = 12 - real_count
        annual_kwh = float(pd.to_numeric(hist_df["Consumo (kWh)"], errors="coerce").fillna(0).sum())
        annual_cost = float(pd.to_numeric(hist_df["Custo (R$)"], errors="coerce").fillna(0).sum())
        c1, c2, c3 = st.columns(3)
        c1.metric("Consumo anual (kWh)", fmt_num(annual_kwh, 0))
        c2.metric("Custo anual (R$)", fmt_num(annual_cost, 2))
        c3.metric("Dados reais vs estimados", f"{real_count} reais / {est_count} estimados")
        st.bar_chart(hist_df.set_index("Mês")[["Consumo (kWh)"]])
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

        energia_total = float(
            pd.to_numeric(hist_df_a["Energia ponta (kWh)"], errors="coerce").fillna(0).sum()
            + pd.to_numeric(hist_df_a["Energia fora ponta (kWh)"], errors="coerce").fillna(0).sum()
        )
        st.metric("Energia anual total (kWh)", fmt_num(energia_total, 0))
        st.line_chart(
            hist_df_a.set_index("Mês")[["Energia ponta (kWh)", "Energia fora ponta (kWh)"]]
        )

if step == 3:
    st.header("ETAPA 3 — Perfil de Consumo")
    st.info("Agora traduzimos sua conta de energia em comportamento energético.")
    profiles = load_typical_profiles()
    profile_map = {p["nome"]: p for p in profiles}

    tipo = st.selectbox("Perfil típico", options=list(profile_map.keys()))
    dias_func = st.number_input("Dias de funcionamento/mês", min_value=1, max_value=31, value=22)
    peak_start = st.text_input("Início ponta", value="18:00")
    peak_end = st.text_input("Fim ponta", value="21:00")

    group = st.session_state["tariff_group"]
    if group == "A":
        c1, c2 = st.columns(2)
        with c1:
            ep = st.number_input("Energia ponta mensal (kWh)", min_value=0.0, value=1000.0)
            dp = st.number_input("Demanda ponta (kW)", min_value=0.0, value=120.0)
        with c2:
            ef = st.number_input("Energia fora ponta mensal (kWh)", min_value=0.0, value=4000.0)
            dfp = st.number_input("Demanda fora ponta (kW)", min_value=0.0, value=90.0)
        bill = {
            "energia_ponta_kwh_mes": ep,
            "energia_fora_ponta_kwh_mes": ef,
            "demanda_ponta_kw": dp,
            "demanda_fora_ponta_kw": dfp,
        }
    else:
        energy_total = st.number_input("Energia total mensal (kWh)", min_value=0.0, value=5000.0)
        bill = {"energia_total_kwh_mes": energy_total, "demanda_estimada_kw": 0.0}

    if st.button("Calibrar curva e gerar fit"):
        selected = profile_map[tipo]
        peak_labels = classify_peak_hours(list(range(24)), peak_start, peak_end, weekdays_only=False)
        fit = fit_load_profile_to_bill(
            base_profile_24h=selected["curva_24h_pu"],
            bill_data=bill,
            operation_days=int(dias_func),
            peak_hours=peak_labels,
            group_type=group,
        )
        st.session_state["fit_adv"] = fit
        st.session_state["fit_adv_profile"] = selected
        st.success("Curva ajustada com sucesso.")

    fit_adv = st.session_state.get("fit_adv")
    if fit_adv is not None:
        sel = st.session_state["fit_adv_profile"]
        st.line_chart(pd.DataFrame({"Curva original (pu)": sel["curva_24h_pu"], "Curva ajustada (kW)": fit_adv["load_curve_kw"]}))
        warns = fit_adv.get("warnings", [])
        if warns:
            for w in warns:
                st.warning(w)
        else:
            st.success("Fit sem inconsistências críticas.")

if step == 4:
    st.header("ETAPA 4 — Área Disponível")
    st.info("Aqui definimos a limitação física real do projeto.")
    endereco = st.text_input("1) Buscar endereço", value=st.session_state.get("localizacao", "São Paulo, SP"))
    c1, c2 = st.columns(2)
    with c1:
        if st.button("2) Geocodificar endereço"):
            try:
                lat_map, lon_map = geocode_address(endereco)
                st.session_state["adv_map_center"] = (lat_map, lon_map)
            except Exception as e:
                st.error(f"Falha ao geocodificar: {e}")
    with c2:
        pack = st.slider("5) Fator de aproveitamento", 0.50, 0.90, 0.70, 0.01)
        st.session_state["packing_factor"] = pack

    if st.session_state.get("adv_map_center"):
        st.caption("3) Desenhe o polígono no mapa e confirme.")
        map_data = render_area_map(st.session_state["adv_map_center"])
        drawing = map_data.get("last_active_drawing") if isinstance(map_data, dict) else None
        if drawing and drawing.get("geometry", {}).get("type") == "Polygon":
            coords = drawing["geometry"]["coordinates"][0]
            area_m2 = calculate_polygon_area_m2([(c[1], c[0]) for c in coords[:-1]])
            st.session_state["available_area_m2"] = area_m2
            st.success(f"4) Área confirmada: {area_m2:.2f} m²")

    if st.session_state.get("available_area_m2"):
        area = st.session_state["available_area_m2"]
        approx_kwp = area * st.session_state.get("packing_factor", 0.7) / 2.3 * 0.55
        m1, m2 = st.columns(2)
        m1.metric("Área disponível (m²)", fmt_num(area, 2))
        m2.metric("6) Potência máxima estimada (kWp)", fmt_num(approx_kwp, 2))

if step == 5:
    st.header("ETAPA 5 — Dimensionamento Técnico")
    st.info("Nesta etapa, dimensionamos a melhor solução tecnicamente viável.")
    azimuth_manual = st.selectbox("Azimute manual (fallback)", [0, 90, 180, 270], index=2)
    tilt = st.number_input("Inclinação (tilt) [padrão = latitude]", value=float(abs(st.session_state["latitude"])), format="%.2f")
    if st.button("Aplicar tilt padrão da latitude"):
        tilt = float(abs(st.session_state["latitude"]))

    st.caption("Opcional: defina o azimute pelo mapa marcando dois pontos (base da placa e direção da face).")
    if st.session_state.get("adv_map_center"):
        map_orient = render_area_map(st.session_state["adv_map_center"], key="azimuth_map")
        drawing2 = map_orient.get("last_active_drawing") if isinstance(map_orient, dict) else None
        azimuth_map = None
        if drawing2 and drawing2.get("geometry", {}).get("type") == "LineString":
            coords = drawing2["geometry"]["coordinates"]
            if len(coords) >= 2:
                start_lon, start_lat = coords[0]
                end_lon, end_lat = coords[-1]
                azimuth_map = calculate_azimuth_from_points(start_lat, start_lon, end_lat, end_lon)
                st.session_state["azimuth_from_map"] = azimuth_map
                st.success(f"Azimute calculado no mapa: {azimuth_map:.1f}°")
    azimuth = int(round(st.session_state.get("azimuth_from_map", azimuth_manual)))

    hist_kwh_year = 0.0
    if st.session_state["tariff_group"] == "B":
        hist_kwh_year = float(pd.to_numeric(st.session_state["historico_df"]["Consumo (kWh)"], errors="coerce").fillna(0).sum())
    elif "historico_df_a" in st.session_state:
        hda = st.session_state["historico_df_a"]
        hist_kwh_year = float(pd.to_numeric(hda["Energia ponta (kWh)"], errors="coerce").fillna(0).sum() + pd.to_numeric(hda["Energia fora ponta (kWh)"], errors="coerce").fillna(0).sum())

    monthly_target = hist_kwh_year / 12 if hist_kwh_year > 0 else 5000
    st.caption(f"Alvo mensal de energia para dimensionamento: {fmt_num(monthly_target, 1)} kWh")

    if st.button("Executar dimensionamento técnico"):
        resultados_df, erro = realizar_dimensionamento_completo(
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

    res = st.session_state.get("resultados_df_adv")
    if isinstance(res, pd.DataFrame) and not res.empty:
        best = res.iloc[0]
        if best.get("fonte_geracao") == "estimativa_fallback":
            st.warning(
                "PVWatts indisponível no momento. O dimensionamento foi calculado com estimativa offline de produtividade solar."
            )
        area_meta = st.session_state.get("area_meta") or {"module_area_m2": 2.3, "warnings": []}
        area_req = int(best.get("sistema_num_total_paineis", 0)) * float(area_meta.get("module_area_m2", 2.3))
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Potência sugerida (kWp)", fmt_num(float(best.get("sistema_potencia_total_w", 0)) / 1000.0, 2))
        c2.metric("Módulos", fmt_num(best.get("sistema_num_total_paineis", 0), 0))
        c3.metric("Área necessária (m²)", fmt_num(area_req, 2))
        c4.metric("Cobertura energética", "100% alvo")
        with st.expander("Detalhes técnicos e limitações"):
            st.dataframe(res.head(12), use_container_width=True)
            for w in area_meta.get("warnings", []):
                st.warning(w)

if step == 6:
    st.header("ETAPA 6 — Análise Energética Avançada")
    st.info("Agora avaliamos como a geração interage com a operação real do cliente.")
    res = st.session_state.get("resultados_df_adv")
    fit_adv = st.session_state.get("fit_adv")
    if res is None or fit_adv is None:
        st.warning("Para esta etapa, execute antes o fit (Etapa 3) e o dimensionamento (Etapa 5).")
    else:
        best = res.iloc[0]
        pv_hourly = calcular_geracao_horaria_pvwatts(
            potencia_dc_kwp=float(best["sistema_potencia_total_w"]) / 1000.0,
            latitude=float(st.session_state["latitude"]),
            longitude=float(st.session_state["longitude"]),
            azimuth=180,
            tilt=float(abs(st.session_state["latitude"])),
        ) or [0.0] * 8760

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

        k1, k2, k3 = st.columns(3)
        k1.metric("Autoconsumo (%)", fmt_num(sc["self_consumption_ratio"] * 100, 1))
        k2.metric("Injeção na rede (kWh/dia)", fmt_num(sc["total_grid_export_kwh"], 2))
        k3.metric("Peak shaving (kW)", fmt_num(ps["peak_shaving_kw"], 2))
        st.line_chart(pd.DataFrame({"Carga (kW)": load_day, "PV (kW)": pv_day, "Carga líquida (kW)": ps["net_load_kw"]}))

        with st.expander("Painel técnico detalhado"):
            st.json({
                "autoconsumo": sc,
                "peak_shaving": ps,
                "load_shifting": ls,
                "simultaneidade": sc["self_sufficiency_ratio"],
            })

if step == 7:
    st.header("ETAPA 7 — Viabilidade Econômica")
    st.info("Esta etapa traduz engenharia em decisão financeira.")
    adv = st.session_state.get("advanced_metrics")
    res = st.session_state.get("resultados_df_adv")
    if adv is None or res is None:
        st.warning("Execute antes as etapas 5 e 6 para habilitar a análise econômica.")
    else:
        best = res.iloc[0]
        capex = st.number_input("CAPEX estimado (R$)", min_value=1000.0, value=350000.0, step=1000.0)
        tariff = st.number_input("Tarifa média de energia (R$/kWh)", min_value=0.1, value=0.85, step=0.01)

        annual_generation = float(best.get("energia_gerada_anual_kwh", 0.0))
        annual_savings = annual_generation * tariff
        payback = capex / annual_savings if annual_savings > 0 else None
        roi = ((annual_savings * 25 - capex) / capex) * 100 if capex > 0 else None
        irr = irr_bisection(capex, annual_savings, years=20)

        st.session_state["economic_metrics"] = {
            "economia_anual": annual_savings,
            "payback": payback,
            "roi": roi,
            "tir": irr,
            "economia_ponta": adv["shift"].get("savings_rs", 0.0),
            "economia_fora_ponta": max(0.0, annual_savings - adv["shift"].get("savings_rs", 0.0)),
            "sensibilidade_tarifaria": {
                "-10%": annual_generation * tariff * 0.9,
                "base": annual_savings,
                "+10%": annual_generation * tariff * 1.1,
            },
        }
        eco = st.session_state["economic_metrics"]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Economia anual (R$)", fmt_num(eco["economia_anual"], 2))
        c2.metric("Payback (anos)", fmt_num(eco["payback"], 2))
        c3.metric("ROI 25 anos (%)", fmt_num(eco["roi"], 1))
        c4.metric("TIR estimada (%)", fmt_num((eco["tir"] or 0) * 100, 2))
        st.bar_chart(pd.DataFrame(eco["sensibilidade_tarifaria"], index=["Economia anual"]).T)

if step == 8:
    st.header("ETAPA 8 — Relatório Final")
    st.info("Consolidação técnica + executiva em formato premium.")

    report = {
        "1_resumo_executivo": {
            "cliente": st.session_state.get("cliente_nome", "N/D"),
            "contexto": f"{st.session_state.get('cliente_tipo', 'N/D')} | Grupo {st.session_state.get('tariff_group', 'N/D')}",
            "local": st.session_state.get("localizacao", "N/D"),
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
            st.session_state.get("resultados_df_adv").iloc[0].to_dict() if isinstance(st.session_state.get("resultados_df_adv"), pd.DataFrame) and not st.session_state.get("resultados_df_adv").empty else {}
        ),
        "7_analise_energetica": st.session_state.get("advanced_metrics", {}),
        "8_viabilidade_economica": st.session_state.get("economic_metrics", {}),
        "9_limitacoes": (st.session_state.get("area_meta") or {}).get("warnings", []),
        "10_proximos_passos": [
            "Validar curva com medição real de carga (15 min).",
            "Refinar premissas tarifárias da distribuidora local.",
            "Executar visita técnica para engenharia de instalação.",
        ],
    }

    st.subheader("Resumo executivo")
    st.json(report["1_resumo_executivo"])

    with st.expander("Memória de cálculo didática (estilo notebook)", expanded=True):
        st.markdown(
            f"""
### Narrativa técnica
1. Com base nos meses informados, o sistema estimou o consumo anual do cliente.
2. A curva típica selecionada foi ajustada para representar o comportamento energético mais provável.
3. A área útil disponível limita a potência máxima instalável.
4. O sistema fotovoltaico foi dimensionado com dados reais de geração solar e equipamentos cadastrados.
5. O resultado foi convertido em indicadores financeiros para apoiar decisão executiva.

### Alertas e premissas
- Premissa geográfica: lat {fmt_num(st.session_state.get('latitude'), 4)} / lon {fmt_num(st.session_state.get('longitude'), 4)}.
- Grupo tarifário considerado: {st.session_state.get('tariff_group', 'N/D')}.
- Limitações reportadas: {', '.join((st.session_state.get('area_meta') or {}).get('warnings', [])) or 'Sem alertas críticos'}.
"""
        )

    report_json = json.dumps(report, ensure_ascii=False, indent=2, default=str)
    st.download_button("Baixar relatório final (JSON)", report_json, file_name="aurumcalc_relatorio_final.json", mime="application/json")
    zip_bytes = build_latex_report_zip(report)
    st.download_button(
        "Baixar pacote LaTeX + gráficos (.zip)",
        zip_bytes,
        file_name="aurumcalc_relatorio_latex.zip",
        mime="application/zip",
    )

st.markdown("---")
nav_prev, nav_next = st.columns([1, 1])
with nav_prev:
    if st.button("⬅️ Voltar", disabled=st.session_state["current_step"] <= 1):
        st.session_state["current_step"] -= 1
        st.rerun()
with nav_next:
    if st.button("Próximo ➡️", disabled=st.session_state["current_step"] >= 8):
        st.session_state["current_step"] += 1
        st.rerun()
