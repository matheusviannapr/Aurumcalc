# app.py
import os
import sys
import traceback
import inspect
import json
from pathlib import Path
import pandas as pd
import streamlit as st
import numpy as np

# --- Se existir segredo, propaga para variável de ambiente usada no pv_calculator ---
if "PVWATTS_API_KEY" in st.secrets:
    os.environ["PVWATTS_API_KEY"] = st.secrets["PVWATTS_API_KEY"]

# --- Imports do seu módulo de cálculo ---
from pv_calculator import (
    realizar_dimensionamento_completo,
    carregar_dados_equipamentos,
    geocode_location,
    salvar_novo_painel,
    salvar_novo_inversor,
    calcular_geracao_horaria_pvwatts,
    aplicar_restricao_area_resultados,
)
from hourly_analysis import (
    BillingGroupAInput,
    BillingGroupBInput,
    list_profiles_by_sector,
    load_profile,
    calibrate_load_profile,
    classify_tariff_periods,
    month_hour_index,
    calculate_hourly_energy_balance,
    split_by_tariff_period,
    compute_economic_summary,
    simulate_load_shifting,
    simulate_peak_shaving,
    calculate_area_limited_pv,
)
from src.load_profiles import load_typical_profiles
from src.load_profile_fitting import fit_load_profile_to_bill, adjusted_profile_from_fit
from src.tariff_periods import classify_peak_hours
from src.map_area import geocode_address, calculate_polygon_area_m2, render_area_map
from src.self_consumption import analyze_self_consumption
from src.peak_shaving import analyze_peak_shaving
from src.load_shifting import simulate_simplified_load_shifting

# --- Import opcional do gerador de LaTeX da memória de cálculo ---
BUILD_TEX = True
try:
    from memoria_calculo import gerar_memoria_calculo_latex
except Exception:
    BUILD_TEX = False

# =========================================================
# CONFIGURAÇÃO DA PÁGINA
# =========================================================
st.set_page_config(
    page_title="Dimensionamento Fotovoltaico Integrado",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================================
# ESTILO (tema azul marinho com fonte branca)
# =========================================================
custom_css = """
<style>
    .stApp { background-color: #001f3f; color: white; }
    p, h1, h2, h3, h4, h5, h6, .stMarkdown { color: white; }
    .css-1d391kg, .css-1lcbmhc { background-color: #003366; color: white; }
    .stTextInput > div > div > input, .stNumberInput > div > div > input {
        background-color: #004080; color: white;
    }
    .stForm label, .stTextInput label, .stNumberInput label, .stSelectbox label { color: white; }
    .stButton>button { background-color: #007bff; color: white; border-radius: 5px; }
    .stButton>button:hover { background-color: #0056b3; color: white; }
    .dataframe { color: black !important; background-color: white !important; }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# =========================================================
# VARIÁVEIS GLOBAIS
# =========================================================
FILE_PATH_EQUIPAMENTOS = "BDFotovoltaica.xlsx"
LOAD_PROFILE_INDEX = Path("data/load_profiles/profiles_index.json")

# =========================================================
# SESSION STATE (defaults)
# =========================================================
if "latitude" not in st.session_state:
    st.session_state["latitude"] = -20.4600
if "longitude" not in st.session_state:
    st.session_state["longitude"] = -54.6200
if "search_lat" not in st.session_state:
    st.session_state["search_lat"] = None
if "search_lon" not in st.session_state:
    st.session_state["search_lon"] = None
if "fit_result" not in st.session_state:
    st.session_state["fit_result"] = None
if "fit_meta" not in st.session_state:
    st.session_state["fit_meta"] = None

# =========================================================
# FUNÇÕES AUXILIARES
# =========================================================
def fmt_num(x, casas=2, default="0"):
    try:
        s = f"{float(x):,.{casas}f}"
        return s.replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return default

def load_data():
    """Carrega o BD de equipamentos (planilha) com tratamento de erro."""
    try:
        df_paineis, df_inversores = carregar_dados_equipamentos(FILE_PATH_EQUIPAMENTOS)
        return df_paineis, df_inversores
    except Exception as e:
        st.error(f"Erro ao carregar {FILE_PATH_EQUIPAMENTOS}: {e}")
        return pd.DataFrame(), pd.DataFrame()


def get_equipment_session_data():
    """
    Mantém os dados de equipamentos em sessão para evitar perda visual após inclusão.
    """
    if "df_paineis" not in st.session_state or "df_inversores" not in st.session_state:
        df_p, df_i = load_data()
        st.session_state["df_paineis"] = df_p
        st.session_state["df_inversores"] = df_i
    return st.session_state["df_paineis"], st.session_state["df_inversores"]


def refresh_equipment_session_data():
    df_p, df_i = load_data()
    st.session_state["df_paineis"] = df_p
    st.session_state["df_inversores"] = df_i


def add_custom_standard_curve_json(
    profile_name: str,
    setor: str,
    subtipo: str,
    hourly_values_text: str,
):
    """
    Facilita inclusão de curva padrão via sessão/UI.
    Aceita vetor de 24 ou 8760 valores em JSON.
    """
    if not profile_name.strip():
        raise ValueError("Informe um nome para a curva.")

    values = json.loads(hourly_values_text)
    if not isinstance(values, list) or len(values) not in (24, 8760):
        raise ValueError("A curva deve ter 24 ou 8760 pontos.")

    profile_id = f"{setor}_{subtipo}_{profile_name}".lower().replace(" ", "_")
    folder = Path("data/load_profiles") / "custom"
    folder.mkdir(parents=True, exist_ok=True)
    file_path = folder / f"{profile_id}.json"

    payload = {
        "id": profile_id,
        "nome": profile_name,
        "setor": setor,
        "subtipo": subtipo,
        "pais_origem": "BR",
        "fonte": "custom_user_import",
        "resolucao_original": f"{len(values)}h",
        "resolucao_final": f"{len(values)}h",
        "ano_climatico_ou_origem": "custom_uploaded",
        "unidade_original": "normalized",
        "tipo_perfil": "custom_uploaded",
        "observacoes": "Curva importada manualmente pelo usuário.",
        "hourly_vector_normalized": values,
    }
    file_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    index_data = {"profiles": []}
    if LOAD_PROFILE_INDEX.exists():
        index_data = json.loads(LOAD_PROFILE_INDEX.read_text(encoding="utf-8"))
    profiles = index_data.setdefault("profiles", [])
    profiles = [p for p in profiles if p.get("id") != profile_id]
    profiles.append(
        {
            "id": profile_id,
            "setor": setor,
            "path": f"custom/{file_path.name}",
            "tipo_perfil": "custom_uploaded",
            "nome": profile_name,
        }
    )
    index_data["profiles"] = profiles
    LOAD_PROFILE_INDEX.write_text(json.dumps(index_data, ensure_ascii=False, indent=2), encoding="utf-8")

def apply_coordinates():
    """Aplica coordenadas encontradas à UI e ao session_state."""
    if st.session_state["search_lat"] is not None and st.session_state["search_lon"] is not None:
        st.session_state["latitude_input"] = st.session_state["search_lat"]
        st.session_state["longitude_input"] = st.session_state["search_lon"]
        st.session_state["latitude"] = st.session_state["search_lat"]
        st.session_state["longitude"] = st.session_state["search_lon"]
        st.session_state["search_lat"] = None
        st.session_state["search_lon"] = None
        st.success(
            f"Coordenadas aplicadas: Lat={st.session_state['latitude']:.6f}, "
            f"Lon={st.session_state['longitude']:.6f}"
        )
    else:
        st.warning("Nenhuma coordenada para aplicar. Faça a busca primeiro.")

def search_coordinates(location_name: str):
    """Busca coordenadas via geocode_location e guarda no session_state."""
    if location_name:
        with st.spinner(f"Buscando coordenadas para '{location_name}'..."):
            lat_geo, lon_geo = geocode_location(location_name)
            if lat_geo is not None and lon_geo is not None:
                st.session_state["search_lat"] = float(lat_geo)
                st.session_state["search_lon"] = float(lon_geo)
                st.success(
                    f"Localização encontrada: Lat={lat_geo:.6f}, Lon={lon_geo:.6f}. "
                    "Clique em 'Aplicar Coordenadas' para usar."
                )
            else:
                st.error("Não foi possível encontrar coordenadas para a localização informada.")
                st.session_state["search_lat"] = None
                st.session_state["search_lon"] = None
    else:
        st.warning("Digite o nome da localização.")

# =========================================================
# SIDEBAR
# =========================================================
st.sidebar.title("Configurações do Sistema")
st.sidebar.markdown("---")
st.sidebar.markdown("Desenvolvido por **Matheus Vianna**")
st.sidebar.markdown("[matheusvianna.com](https://matheusvianna.com)")
st.sidebar.markdown("---")

df_paineis, df_inversores = get_equipment_session_data()

with st.sidebar.expander("Dados de Equipamentos Carregados"):
    st.subheader("Painéis Solares")
    if not df_paineis.empty:
        st.dataframe(
            df_paineis[["modelo", "potencia_maxima_nominal_pmax", "tensao_circuito_aberto_voc"]].head()
        )
    else:
        st.info("Base de painéis vazia.")
    st.subheader("Inversores")
    if not df_inversores.empty:
        st.dataframe(
            df_inversores[["modelo", "maxima_potencia_nominal_ca", "tensao_maxima_cc"]].head()
        )
    else:
        st.info("Base de inversores vazia.")

with st.sidebar.expander("➕ Inserir Novo Equipamento"):
    st.subheader("Novo Painel Solar")
    with st.form("form_novo_painel", clear_on_submit=True):
        modelo_p = st.text_input("Modelo", key="modelo_p")
        fabricante_p = st.text_input("Fabricante", key="fabricante_p")
        pmax = st.number_input("Potência Máxima Nominal (Pmax) [Wp]", min_value=1.0, step=1.0, format="%.2f", key="pmax_p")
        voc = st.number_input("Tensão de Circuito Aberto (Voc) [V]", min_value=1.0, step=0.1, format="%.2f", key="voc_p")
        isc = st.number_input("Corrente de Curto Circuito (Isc) [A]", min_value=0.1, step=0.1, format="%.2f", key="isc_p")
        vmp = st.number_input("Tensão de Operação Ótima (Vmp) [V]", min_value=0.1, step=0.1, format="%.2f", key="vmp_p")
        imp = st.number_input("Corrente de Operação Ótima (Imp) [A]", min_value=0.1, step=0.1, format="%.2f", key="imp_p")
        eficiencia = st.number_input("Eficiência do Módulo [%]", min_value=1.0, max_value=100.0, step=0.1, format="%.2f", key="eficiencia_p")
        submitted_painel = st.form_submit_button("Salvar Painel")
        if submitted_painel:
            if modelo_p and pmax > 0 and voc > 0 and isc > 0:
                novo_painel_data = {
                    "modelo": modelo_p,
                    "fabricante": fabricante_p,
                    "potencia_maxima_nominal_pmax": pmax,
                    "tensao_operacao_otima_vmp": vmp,
                    "corrente_operacao_otima_imp": imp,
                    "tensao_circuito_aberto_voc": voc,
                    "corrente_curto_circuito_isc": isc,
                    "eficiencia_modulo": eficiencia
                }
                if salvar_novo_painel(novo_painel_data):
                    refresh_equipment_session_data()
                    st.success(f"Painel '{modelo_p}' salvo com sucesso e já disponível na sessão.")
                else:
                    st.error("Erro ao salvar o painel no arquivo Excel.")
            else:
                st.error("Preencha os campos obrigatórios (Modelo, Pmax, Voc, Isc).")

    st.markdown("---")
    st.subheader("Novo Inversor")
    with st.form("form_novo_inversor", clear_on_submit=True):
        modelo_i = st.text_input("Modelo", key="modelo_i")
        fabricante_i = st.text_input("Fabricante", key="fabricante_i")
        pot_ca = st.number_input("Potência Nominal CA Máxima [W]", min_value=1.0, step=1.0, format="%.2f", key="pot_ca_i")
        vmax_cc = st.number_input("Tensão Máxima CC [V]", min_value=1.0, step=1.0, format="%.0f", key="vmax_cc_i")
        vstart = st.number_input("Tensão de Start [V]", min_value=1.0, step=1.0, format="%.0f", key="vstart_i")
        imax_mppt = st.number_input("Corrente Máxima de Entrada por MPPT [A]", min_value=0.1, step=0.1, format="%.2f", key="imax_mppt_i")
        num_mppt = st.number_input("Número de MPPT Trackers", min_value=1, step=1, format="%d", key="num_mppt_i")
        pot_fv_max = st.number_input("Potência Máxima FV (Máxima) [W]", min_value=1.0, step=1.0, format="%.2f", key="pot_fv_max_i")
        v_nom = st.number_input("Tensão Nominal [V]", min_value=1.0, step=1.0, format="%.0f", key="v_nom_i")
        faixa_mpp = st.text_input("Faixa de Tensão MPPT (Ex: 60V-550V)", key="faixa_mpp_i")
        isc_mppt = st.number_input("Corrente Máxima Curto Circuito por MPPT [A]", min_value=0.1, step=0.1, format="%.2f", key="isc_mppt_i")
        pot_ap_ca = st.number_input("Potência Máxima Aparente CA [VA]", min_value=1.0, step=1.0, format="%.2f", key="pot_ap_ca_i")
        v_nom_ca = st.number_input("Tensão Nominal CA [V]", min_value=1.0, step=1.0, format="%.0f", key="v_nom_ca_i")
        freq_ca = st.text_input("Frequência da Rede CA (Ex: 50Hz/60Hz)", key="freq_ca_i")
        i_saida_max = st.number_input("Corrente de Saída Máxima [A]", min_value=0.1, step=0.1, format="%.2f", key="i_saida_max_i")
        fp_ajustavel = st.text_input("Fator de Potência Ajustável (Ex: 0.8i-0.8c)", key="fp_ajustavel_i")
        fases_ca = st.number_input("Quantidade de Fases CA", min_value=1, step=1, format="%d", key="fases_ca_i")

        submitted_inversor = st.form_submit_button("Salvar Inversor")
        if submitted_inversor:
            if (modelo_i and pot_ca > 0 and vmax_cc > 0 and vstart > 0 and
                imax_mppt > 0 and num_mppt > 0):
                novo_inversor_data = {
                    "modelo": modelo_i,
                    "fabricante": fabricante_i,
                    "potencia_maxima_fv_maxima": pot_fv_max,
                    "tensao_maxima_cc": vmax_cc,
                    "tensao_start": vstart,
                    "tensao_nominal": v_nom,
                    "faixa_tensao_mpp": faixa_mpp,
                    "numero_mpp_trackers": num_mppt,
                    "corrente_maxima_entrada_por_mppt_tracker": imax_mppt,
                    "corrente_maxima_curto_circuito_por_mppt_tracker": isc_mppt,
                    "maxima_potencia_nominal_ca": pot_ca,
                    "potencia_maxima_aparente_ca": pot_ap_ca,
                    "tensao_nominal_ca": v_nom_ca,
                    "frequencia_rede_ca": freq_ca,
                    "corrente_saida_maxima": i_saida_max,
                    "fator_potencia_ajustavel": fp_ajustavel,
                    "quantidade_fases_ca": fases_ca
                }
                if salvar_novo_inversor(novo_inversor_data):
                    refresh_equipment_session_data()
                    st.success(f"Inversor '{modelo_i}' salvo com sucesso e já disponível na sessão.")
                else:
                    st.error("Erro ao salvar o inversor no arquivo Excel.")
            else:
                st.error("Preencha os campos obrigatórios (Modelo, Potência CA, Vmax CC, Vstart, Imax MPPT, Num MPPT).")

# =========================================================
# TÍTULO
# =========================================================
st.title("☀️ Dimensionamento Fotovoltaico Integrado")
st.markdown("---")

# =========================================================
# INPUTS
# =========================================================
st.header("1. Dados de Geração (PVWatts)")
col1, col2 = st.columns(2)

with col1:
    consumo_medio_mensal = st.number_input(
        "Consumo Médio Mensal (kWh)",
        min_value=1,
        value=300,
        step=10,
        help="Consumo médio mensal em kWh."
    )
    location_name = st.text_input(
        "Pesquisar Localização por Nome (Ex: São Paulo, SP)",
        help="Digite o nome da cidade ou endereço para obter as coordenadas."
    )
    if st.button("Buscar Coordenadas"):
        search_coordinates(location_name)

if st.session_state["search_lat"] is not None and st.session_state["search_lon"] is not None:
    st.button(
        f"Aplicar Coordenadas Encontradas: Lat={st.session_state['search_lat']:.6f}, "
        f"Lon={st.session_state['search_lon']:.6f}",
        on_click=apply_coordinates
    )

with col2:
    latitude = st.number_input(
        "Latitude (°)",
        value=float(st.session_state["latitude"]),
        format="%.6f",
        key="latitude_input",
        help="Latitude do local de instalação."
    )
    longitude = st.number_input(
        "Longitude (°)",
        value=float(st.session_state["longitude"]),
        format="%.6f",
        key="longitude_input",
        help="Longitude do local de instalação."
    )
    st.session_state["latitude"] = latitude
    st.session_state["longitude"] = longitude

col3, col4 = st.columns(2)
with col3:
    azimuth = st.selectbox(
        "Azimuth (°)",
        options=[0, 90, 180, 270],
        index=2,  # 180 padrão
        help="0°=Norte, 90°=Leste, 180°=Sul, 270°=Oeste."
    )
with col4:
    tilt_sugerido = abs(latitude)
    tilt = st.number_input(
        f"Tilt (Inclinação) Sugerido: {tilt_sugerido:.2f}°",
        value=float(tilt_sugerido),
        format="%.2f",
        help="Inclinação dos painéis (sugestão = latitude)."
    )

st.markdown("---")

# =========================================================
# DIAGNÓSTICO OPCIONAL (ver assinatura da função)
# =========================================================
with st.expander("🔧 Diagnóstico (opcional)"):
    try:
        st.write("Assinatura realizar_dimensionamento_completo:",
                 str(inspect.signature(realizar_dimensionamento_completo)))
    except Exception as e:
        st.write("Não foi possível inspecionar a assinatura:", e)

# =========================================================
# BOTÃO: CALCULAR
# =========================================================
if st.button("Realizar Dimensionamento Completo"):
    if consumo_medio_mensal <= 0:
        st.error("O consumo médio mensal deve ser maior que zero.")
        st.stop()

    # Sanidade de coordenadas (evita None e strings)
    try:
        lat_val = float(latitude)
        lon_val = float(longitude)
    except Exception:
        st.error("Coordenadas inválidas. Verifique Latitude e Longitude.")
        st.stop()

    with st.spinner("Calculando potência de pico e dimensionando o sistema..."):
        try:
            # Assinatura esperada:
            # realizar_dimensionamento_completo(consumo_medio_mensal, latitude, longitude, azimuth, tilt)
            resultados_df, erro = realizar_dimensionamento_completo(
                consumo_medio_mensal, lat_val, lon_val, int(azimuth), float(tilt)
            )
        except Exception as e:
            st.error(f"{type(e).__name__}: {e}")
            st.code("".join(traceback.format_exception(*sys.exc_info())))
            st.stop()

        if erro:
            st.error(f"Ocorreu um erro durante o dimensionamento: {erro}")
            st.stop()

        if resultados_df is None or resultados_df.empty:
            st.error("Nenhum resultado foi retornado.")
            st.stop()

        st.success("Dimensionamento realizado com sucesso!")

        # =========================================================
        # RESULTADOS
        # =========================================================
        st.header("2. Resultados do Dimensionamento")
        st.subheader("2.1. Resumo de Geração")

        # Pega com fallback seguro
        try:
            potencia_pico_necessaria_kw = resultados_df.get("potencia_pico_necessaria_kw", pd.Series([None])).iloc[0]
        except Exception:
            potencia_pico_necessaria_kw = None

        try:
            consumo_anual_kwh = resultados_df.get("consumo_anual_kwh", pd.Series([consumo_medio_mensal * 12])).iloc[0]
        except Exception:
            consumo_anual_kwh = consumo_medio_mensal * 12

        try:
            energia_gerada_anual_kwh = resultados_df.get("energia_gerada_anual_kwh", pd.Series([None])).iloc[0]
        except Exception:
            energia_gerada_anual_kwh = None

        col_resumo1, col_resumo2, col_resumo3 = st.columns(3)
        with col_resumo1:
            st.metric("Consumo Anual Alvo (kWh)", fmt_num(consumo_anual_kwh, 2, default="-"))
        with col_resumo2:
            st.metric("Potência de Pico Necessária (kWp)", fmt_num(potencia_pico_necessaria_kw, 2, default="-"))
        with col_resumo3:
            st.metric("Energia Anual Estimada (kWh)", fmt_num(energia_gerada_anual_kwh, 2, default="-"))

        st.markdown("---")

        # TABELA DE OPÇÕES
        st.subheader("2.2. Opções de Dimensionamento (Inversor e Arranjo)")
        colunas_esperadas = [
            "inversor_modelo", "inversor_fabricante", "inversor_num_unidades",
            "painel_modelo", "painel_fabricante", "painel_potencia",
            "arranjo_modulos_serie", "arranjo_conjuntos_paralelo_por_mppt",
            "inversor_num_mppt", "sistema_num_total_paineis", "sistema_potencia_total_w"
        ]
        cols_existentes = [c for c in colunas_esperadas if c in resultados_df.columns]
        df_display = resultados_df[cols_existentes].copy()

        renomear = {
            "inversor_modelo": "Inversor (Modelo)",
            "inversor_fabricante": "Inversor (Fabricante)",
            "inversor_num_unidades": "Inversor (Qtd.)",
            "painel_modelo": "Painel (Modelo)",
            "painel_fabricante": "Painel (Fabricante)",
            "painel_potencia": "Painel (Potência Wp)",
            "arranjo_modulos_serie": "Módulos em Série (por MPPT)",
            "arranjo_conjuntos_paralelo_por_mppt": "Conjuntos em Paralelo (por MPPT)",
            "inversor_num_mppt": "MPPTs (por Inversor)",
            "sistema_num_total_paineis": "Total de Painéis",
            "sistema_potencia_total_w": "Potência Total do Sistema (Wp)"
        }
        df_display.rename(columns=renomear, inplace=True)

        if "Painel (Potência Wp)" in df_display.columns:
            df_display["Painel (Potência Wp)"] = df_display["Painel (Potência Wp)"].apply(
                lambda x: fmt_num(x, 0, default="-")
            )
        if "Potência Total do Sistema (Wp)" in df_display.columns:
            df_display["Potência Total do Sistema (Wp)"] = df_display["Potência Total do Sistema (Wp)"].apply(
                lambda x: fmt_num(x, 0, default="-")
            )

        st.dataframe(df_display, use_container_width=True)

        st.markdown("""
        <div style='color: white; font-size: small;'>
        <strong>Explicação da Tabela:</strong> Cada linha representa uma opção de dimensionamento válida. 
        A coluna 'Potência Total do Sistema (Wp)' indica a potência real instalada, que deve ser próxima da 'Potência de Pico Necessária'.
        O arranjo é detalhado por MPPT (Maximum Power Point Tracker) do inversor.
        </div>
        """, unsafe_allow_html=True)

        st.markdown("---")

        # DETALHES TÉCNICOS
        st.subheader("2.3. Detalhes Técnicos do Melhor Arranjo")

        def fmt_int(x):
            try:
                return fmt_num(x, 0, default="-")
            except Exception:
                return str(x)

        melhor_arranjo = resultados_df.iloc[0]

        st.markdown(f"""
        <div style='background-color: #004080; padding: 15px; border-radius: 10px;'>
        <strong>Melhor Opção Selecionada:</strong><br><br>
        - <strong>Inversor:</strong> {melhor_arranjo.get('inversor_modelo', '')} ({melhor_arranjo.get('inversor_fabricante','')})<br>
        - <strong>Quantidade de Inversores:</strong> {melhor_arranjo.get('inversor_num_unidades','')}<br>
        - <strong>Painel:</strong> {melhor_arranjo.get('painel_modelo','')} ({melhor_arranjo.get('painel_fabricante','')}) - {fmt_int(melhor_arranjo.get('painel_potencia',0))} Wp<br><br>
        <strong>Detalhes do Arranjo (por MPPT):</strong><br>
        - <strong>Módulos em Série:</strong> {melhor_arranjo.get('arranjo_modulos_serie','')}<br>
        - <strong>Conjuntos em Paralelo:</strong> {melhor_arranjo.get('arranjo_conjuntos_paralelo_por_mppt','')}<br>
        - <strong>Potência do Arranjo (por MPPT):</strong> {fmt_int(melhor_arranjo.get('arranjo_potencia_total_mppt_w',0))} Wp<br><br>
        <strong>Sistema Total:</strong><br>
        - <strong>Potência Total Instalada:</strong> {fmt_int(melhor_arranjo.get('sistema_potencia_total_w',0))} Wp<br>
        - <strong>Total de Painéis:</strong> {melhor_arranjo.get('sistema_num_total_paineis','')}<br>
        </div>
        """, unsafe_allow_html=True)

        # Dados brutos (opcional)
        with st.expander("Ver Dados Brutos do Dimensionamento"):
            st.dataframe(resultados_df, use_container_width=True)

        # =========================================================
        # MEMÓRIA DE CÁLCULO (LaTeX) — agora lendo o arquivo gerado
        # =========================================================
        if BUILD_TEX:
            try:
                ac_monthly = melhor_arranjo.get("energia_mensal_kwh_array")

                # Monta um payload tolerante (a função aceita DataFrame/dict/list[dict])
                payload = {
                    "projeto": "Dimensionamento FV — IME",
                    "cliente": "Instituto Militar de Engenharia",
                    "local": "Rio de Janeiro, RJ",
                    "consumo_anual_kwh": consumo_medio_mensal * 12,
                    "potencia_pico_necessaria_kw": melhor_arranjo.get("potencia_pico_necessaria_kw"),
                    "energia_gerada_anual_kwh": melhor_arranjo.get("energia_gerada_anual_kwh"),
                    "energia_mensal_kwh_array": ac_monthly,  # pode ser None
                    "inversor_modelo": melhor_arranjo.get("inversor_modelo"),
                    "inversor_fabricante": melhor_arranjo.get("inversor_fabricante"),
                    "inversor_num_unidades": melhor_arranjo.get("inversor_num_unidades"),
                    "inversor_num_mppt": melhor_arranjo.get("inversor_num_mppt"),
                    "painel_modelo": melhor_arranjo.get("painel_modelo"),
                    "painel_fabricante": melhor_arranjo.get("painel_fabricante"),
                    "painel_potencia": melhor_arranjo.get("painel_potencia"),
                    "arranjo_modulos_serie": melhor_arranjo.get("arranjo_modulos_serie"),
                    "arranjo_conjuntos_paralelo_por_mppt": melhor_arranjo.get("arranjo_conjuntos_paralelo_por_mppt"),
                    "sistema_num_total_paineis": melhor_arranjo.get("sistema_num_total_paineis"),
                    "sistema_potencia_total_w": melhor_arranjo.get("sistema_potencia_total_w"),
                    "consumo_mensal_kwh": consumo_medio_mensal,
                    # meta-dados de sítio
                    "latitude": lat_val,
                    "longitude": lon_val,
                    "azimuth": int(azimuth),
                    "tilt": float(tilt),
                }

                # 1) Gera o arquivo .tex e recebe o CAMINHO
                caminho_tex = gerar_memoria_calculo_latex(
                    resultados_df=payload,
                    caminho_tex="memoria_calculo.tex",
                    projeto="Dimensionamento FV — IME",
                    cliente="Instituto Militar de Engenharia",
                    local="Rio de Janeiro, RJ",
                    latitude=lat_val,
                    longitude=lon_val,
                    azimuth=int(azimuth),
                    tilt=float(tilt),
                    observacoes="Documento gerado automaticamente pelo aplicativo.",
                )

                # 2) Lê o conteúdo do .tex para enviar no download_button
                try:
                    with open(caminho_tex, "r", encoding="utf-8") as f:
                        tex_str = f.read()
                except Exception as e_read:
                    # Fallback: se por algum motivo não conseguiu ler, mostra aviso
                    st.warning(f"Memória gerada, mas não foi possível ler o arquivo: {e_read}")
                    tex_str = ""

                if tex_str:
                    

                    st.download_button(
                        label="⬇️ Baixar Memória de Cálculo (.tex)",
                        data=tex_str,
                        file_name="memoria_calculo.tex",
                        mime="application/x-latex",
                    )
                else:
                    st.warning("Memória de cálculo foi gerada, mas o conteúdo está vazio.")

            except Exception as e:
                st.warning(f"Não foi possível gerar a memória de cálculo LaTeX: {e}")
        else:
            st.info("Para habilitar a Memória de Cálculo (LaTeX), adicione o arquivo 'memoria_calculo.py' ao projeto.")

# =========================================================
# EXPLICAÇÃO
# =========================================================
st.markdown("---")
st.header("Como Funciona o Dimensionamento?")
st.markdown("""
Este aplicativo integra duas etapas cruciais do dimensionamento fotovoltaico:

1. **Cálculo de Geração (PVWatts):**
   - Usa a API PVWatts (NREL) para estimar a produção de energia.
   - A partir de **Consumo Médio Mensal** e **Latitude/Longitude/Azimuth/Tilt**, calcula a **Potência de Pico Necessária (kWp)**.

2. **Seleção de Inversor e Arranjo:**
   - Com a potência alvo, consulta o **banco de dados de Painéis e Inversores**.
   - Define combinações de inversores e arranjos série/paralelo por MPPT respeitando limites elétricos.
   - Retorna opções ordenadas, priorizando a mais próxima da potência alvo.
""")

st.markdown("---")
st.header("Análise Horária e Tarifária")
st.caption("Versão inicial com perfis synthetic_for_testing. Não usar para proposta real sem validação de curva medida.")

col_h1, col_h2 = st.columns(2)
with col_h1:
    customer_sector = st.selectbox(
        "Tipo de cliente",
        options=[
            "escola", "supermercado", "frigorifico", "industria", "escritorio",
            "varejo", "restaurante", "hotel", "hospital", "custom"
        ],
        index=0
    )
    dias_operacao = st.slider("Dias de funcionamento por semana", 1, 7, 5)
    horario_operacao = st.slider("Horário de funcionamento", 0, 24, (8, 18))
with col_h2:
    grupo = st.radio("Grupo tarifário", options=["A", "B"], horizontal=True)
    mes_ref = st.text_input("Mês de referência (YYYY-MM)", value="2026-01")
    dias_ciclo = st.number_input("Dias do ciclo", min_value=1, max_value=31, value=30)
    ponta_window = st.slider("Janela de ponta (hora início/fim)", 0, 24, (18, 21))

profiles = list_profiles_by_sector(customer_sector)
profile_ids = [p["id"] for p in profiles]
profile_choice = st.selectbox("Perfil base", options=profile_ids if profile_ids else ["sem_perfil"])

with st.expander("➕ Incluir curva padrão personalizada (sessão de perfis)"):
    st.caption("Cole um JSON com 24 ou 8760 valores normalizados. Exemplo: [0.04, 0.04, ...]")
    with st.form("form_add_curve", clear_on_submit=True):
        curve_name = st.text_input("Nome da curva")
        curve_setor = st.selectbox(
            "Setor da curva",
            options=[
                "escola", "supermercado", "frigorifico", "industria", "escritorio",
                "varejo", "restaurante", "hotel", "hospital", "custom"
            ],
            index=9
        )
        curve_subtipo = st.text_input("Subtipo", value="custom")
        curve_json = st.text_area(
            "Vetor horário (JSON)",
            value="[0.0416667, 0.0416667, 0.0416667, 0.0416667, 0.0416667, 0.0416667, 0.0416667, 0.0416667, 0.0416667, 0.0416667, 0.0416667, 0.0416667, 0.0416667, 0.0416667, 0.0416667, 0.0416667, 0.0416667, 0.0416667, 0.0416667, 0.0416667, 0.0416667, 0.0416667, 0.0416667, 0.0416667]",
            height=120,
        )
        submitted_curve = st.form_submit_button("Salvar curva padrão")
        if submitted_curve:
            try:
                add_custom_standard_curve_json(curve_name, curve_setor, curve_subtipo, curve_json)
                st.success("Curva padrão salva com sucesso. Reabra o seletor para visualizar o novo perfil.")
                st.rerun()
            except Exception as e:
                st.error(f"Não foi possível salvar a curva: {e}")

if grupo == "A":
    c1, c2, c3 = st.columns(3)
    with c1:
        energia_ponta = st.number_input("Energia ponta (kWh)", min_value=0.0, value=1000.0, step=10.0)
        demanda_ponta = st.number_input("Demanda ponta (kW)", min_value=0.0, value=120.0, step=1.0)
        tarifa_ponta = st.number_input("Tarifa energia ponta (R$/kWh)", min_value=0.0, value=1.2, step=0.01)
    with c2:
        energia_fora = st.number_input("Energia fora ponta (kWh)", min_value=0.0, value=4000.0, step=10.0)
        demanda_fora = st.number_input("Demanda fora ponta (kW)", min_value=0.0, value=90.0, step=1.0)
        tarifa_fora = st.number_input("Tarifa energia fora ponta (R$/kWh)", min_value=0.0, value=0.7, step=0.01)
    with c3:
        tarifa_dem_p = st.number_input("Tarifa demanda ponta (R$/kW)", min_value=0.0, value=30.0, step=0.1)
        tarifa_dem_f = st.number_input("Tarifa demanda fora ponta (R$/kW)", min_value=0.0, value=20.0, step=0.1)

    billing_obj = BillingGroupAInput(
        concessionaria="N/A", modalidade_tarifaria="verde", subgrupo="A4", mes_referencia=mes_ref,
        dias_ciclo=int(dias_ciclo), ponta_inicio_hora=ponta_window[0], ponta_fim_hora=ponta_window[1],
        feriados_nacionais=[], energia_ponta_kwh=energia_ponta, energia_fora_ponta_kwh=energia_fora,
        demanda_medida_ponta_kw=demanda_ponta, demanda_medida_fora_ponta_kw=demanda_fora,
        tarifa_energia_ponta=tarifa_ponta, tarifa_energia_fora_ponta=tarifa_fora,
        tarifa_demanda_ponta=tarifa_dem_p, tarifa_demanda_fora_ponta=tarifa_dem_f,
    )
else:
    c1, c2 = st.columns(2)
    with c1:
        energia_total_b = st.number_input("Energia total (kWh)", min_value=0.0, value=5000.0, step=10.0)
        energia_ponta_b = st.number_input("Energia ponta (branca, opcional)", min_value=0.0, value=0.0, step=10.0)
        energia_fora_b = st.number_input("Energia fora ponta (branca, opcional)", min_value=0.0, value=0.0, step=10.0)
    with c2:
        tarifa_energia_b = st.number_input("Tarifa energia (R$/kWh)", min_value=0.0, value=0.8, step=0.01)
        tarifa_ponta_b = st.number_input("Tarifa ponta (branca)", min_value=0.0, value=1.3, step=0.01)
        tarifa_fora_b = st.number_input("Tarifa fora ponta (branca)", min_value=0.0, value=0.6, step=0.01)
    billing_obj = BillingGroupBInput(
        concessionaria="N/A", modalidade="convencional", mes_referencia=mes_ref, dias_ciclo=int(dias_ciclo),
        energia_total_kwh=energia_total_b, energia_ponta_kwh=energia_ponta_b or None, energia_fora_ponta_kwh=energia_fora_b or None,
        tarifa_energia=tarifa_energia_b, tarifa_energia_ponta=tarifa_ponta_b, tarifa_energia_fora_ponta=tarifa_fora_b,
    )

col_fit, col_calc = st.columns(2)
with col_fit:
    run_fit = st.button("1) Calibrar curva (FIT)")
with col_calc:
    run_sizing = st.button("2) Calcular sistema com base no FIT + conta")

if run_fit:
    try:
        selected_meta = next((p for p in profiles if p["id"] == profile_choice), None)
        if not selected_meta:
            st.error("Nenhum perfil disponível para o setor.")
        else:
            candidate = load_profile(selected_meta["path"])
            calibration = calibrate_load_profile(
                candidate_profiles=[candidate],
                billing_data=billing_obj,
                customer_type=customer_sector,
                dias_funcionamento_semana=dias_operacao,
                horario_funcionamento=horario_operacao,
            )
            st.session_state["fit_result"] = calibration
            st.session_state["fit_meta"] = {
                "mes_ref": mes_ref,
                "dias_ciclo": int(dias_ciclo),
                "ponta_window": ponta_window,
                "grupo": grupo,
            }
            st.success(f"FIT concluído: {calibration.perfil_escolhido_id} | fator: {calibration.fator_aplicado:.4f}")
            if calibration.alertas:
                for alert in calibration.alertas:
                    st.warning(alert)
    except Exception as e:
        st.error(f"Falha na calibração (FIT): {e}")

fit_result = st.session_state.get("fit_result")
fit_meta = st.session_state.get("fit_meta")

if fit_result is not None:
    st.info("FIT disponível em sessão. Agora você pode calcular o sistema com base na conta + FIT.")
    fit_chart = pd.DataFrame({"carga_fit_kw": fit_result.curva_calibrada_kw[:168]})
    st.line_chart(fit_chart)

if run_sizing:
    try:
        if fit_result is None or fit_meta is None:
            st.error("Execute primeiro o passo 1 (Calibrar curva / FIT).")
        else:
            idx = month_hour_index(fit_meta["mes_ref"], int(fit_meta["dias_ciclo"]))
            periods = classify_tariff_periods(idx, fit_meta["ponta_window"][0], fit_meta["ponta_window"][1], [])

            target_monthly_kwh = float(np.sum(fit_result.curva_calibrada_kw))
            if isinstance(billing_obj, BillingGroupAInput):
                billed_total = (billing_obj.energia_ponta_kwh or 0.0) + (billing_obj.energia_fora_ponta_kwh or 0.0)
            else:
                billed_total = billing_obj.energia_total_kwh or target_monthly_kwh
            target_monthly_kwh = billed_total if billed_total > 0 else target_monthly_kwh

            pv_per_kwp = calcular_geracao_horaria_pvwatts(
                potencia_dc_kwp=1.0, latitude=latitude, longitude=longitude, azimuth=int(azimuth), tilt=float(tilt)
            ) or [0.0] * len(idx)
            pv_per_kwp_month = float(np.sum(np.asarray(pv_per_kwp[: len(idx)], dtype=float)))

            if pv_per_kwp_month <= 0:
                st.error("Não foi possível obter geração horária PVWatts para calcular o tamanho do sistema.")
            else:
                suggested_kwp = target_monthly_kwh / pv_per_kwp_month
                sizing_factor = st.slider("Fator do sistema sobre alvo (para autoconsumo/limites)", 0.5, 1.3, 1.0, 0.05)
                system_kwp = max(0.1, suggested_kwp * sizing_factor)

                pv_hourly = calcular_geracao_horaria_pvwatts(
                    potencia_dc_kwp=system_kwp, latitude=latitude, longitude=longitude, azimuth=int(azimuth), tilt=float(tilt)
                ) or [0.0] * len(idx)
                pv_slice = np.asarray(pv_hourly[: len(idx)], dtype=float)
                balance = calculate_hourly_energy_balance(fit_result.curva_calibrada_kw, pv_slice)
                split = split_by_tariff_period(balance, periods)
                econ = compute_economic_summary(split, billing_obj)

                st.subheader("Sistema calculado a partir do FIT + conta")
                c1, c2, c3 = st.columns(3)
                c1.metric("Potência sugerida (kWp)", fmt_num(system_kwp, 2))
                c2.metric("Carga mensal alvo (kWh)", fmt_num(target_monthly_kwh, 2))
                c3.metric("Geração mensal estimada (kWh)", fmt_num(float(np.sum(pv_slice)), 2))

                chart_df = pd.DataFrame({
                    "carga_fit_kw": fit_result.curva_calibrada_kw[:168],
                    "pv_kw": pv_slice[:168],
                    "import_kw": balance["grid_import_kw"].values[:168],
                })
                st.line_chart(chart_df)

                st.subheader("Resumo de economia mensal")
                e1, e2, e3 = st.columns(3)
                e1.metric("Economia ponta (R$)", fmt_num(econ["economia_energia_ponta"], 2))
                e2.metric("Economia fora ponta (R$)", fmt_num(econ["economia_energia_fora_ponta"], 2))
                e3.metric("Economia total mensal (R$)", fmt_num(econ["economia_total_mensal"], 2))

                st.subheader("Área máxima disponível")
                area_limit = calculate_area_limited_pv(
                    area_disponivel_m2=100.0, modulo_area_m2=2.3, fator_ocupacao=0.75,
                    potencia_modulo_w=550.0, potencia_sistema_kwp=system_kwp,
                )
                st.json(area_limit)

                ls = simulate_load_shifting(
                    fit_result.curva_calibrada_kw, periods, percentual_flexivel=0.15,
                    limite_energia_deslocavel_dia_kwh=30.0, limite_potencia_deslocavel_kw=10.0,
                )
                ps = simulate_peak_shaving(
                    balance["grid_import_kw"].values, demand_target_kw=float(np.percentile(balance["grid_import_kw"], 95)),
                    battery_power_kw=20.0, battery_capacity_kwh=60.0,
                )
                with st.expander("Resultados load shifting e peak shaving"):
                    st.write("Load shifting (energia antes/depois):", ls["energia_total_antes"], ls["energia_total_depois"])
                    st.write("Peak shaving demanda antes/depois:", ps["demanda_antes"], ps["demanda_depois"])
                    if ps.get("alertas"):
                        st.warning(ps["alertas"][0])
    except Exception as e:
        st.error(f"Falha no cálculo do sistema a partir do FIT: {e}")

st.markdown('---')
st.header('Dimensionamento com Curva de Demanda Ajustada')
st.warning('As curvas de demanda utilizadas nesta versão são curvas típicas/sintéticas. Para estudos executivos, recomenda-se calibração com memória de massa, medição real ou analisador de energia.')

profiles_typical = load_typical_profiles()
profile_map = {p['nome']: p for p in profiles_typical}

cadv1, cadv2 = st.columns(2)
with cadv1:
    tipo_cliente_nome = st.selectbox('Tipo de cliente (curva típica)', options=list(profile_map.keys()))
    grupo_adv = st.radio('Grupo tarifário (novo modo)', options=['A', 'B'], horizontal=True, key='grupo_adv')
    dias_func_mes = st.number_input('Dias de funcionamento no mês', min_value=1, max_value=31, value=22)
    peak_start = st.text_input('Horário ponta início', value='18:00')
    peak_end = st.text_input('Horário ponta fim', value='21:00')
    weekdays_only_peak = st.checkbox('Considerar ponta apenas em dias úteis', value=False)
with cadv2:
    endereco_mapa = st.text_input('Endereço ou coordenadas para mapa', value='São Paulo, SP')
    packing_factor = st.slider('Fator de aproveitamento da área', min_value=0.50, max_value=0.90, value=0.70, step=0.01)
    pct_shift = st.slider('Percentual de carga deslocável na ponta', min_value=0, max_value=100, value=0, step=5)
    shift_window = st.selectbox('Janela de deslocamento', options=['madrugada', 'horário solar'])

if grupo_adv == 'A':
    a1, a2 = st.columns(2)
    with a1:
        energia_ponta_mes = st.number_input('Energia ponta (kWh/mês)', min_value=0.0, value=1000.0, step=10.0, key='adv_ep')
        demanda_ponta_kw = st.number_input('Demanda ponta (kW)', min_value=0.0, value=120.0, step=1.0, key='adv_dp')
        tarifa_energia_ponta_adv = st.number_input('Tarifa energia ponta (opcional)', min_value=0.0, value=1.2, step=0.01, key='adv_tp')
        tarifa_demanda_ponta_adv = st.number_input('Tarifa demanda ponta (opcional)', min_value=0.0, value=0.0, step=0.01, key='adv_tdp')
    with a2:
        energia_fora_mes = st.number_input('Energia fora ponta (kWh/mês)', min_value=0.0, value=4000.0, step=10.0, key='adv_ef')
        demanda_fora_kw = st.number_input('Demanda fora ponta (kW)', min_value=0.0, value=90.0, step=1.0, key='adv_df')
        tarifa_energia_fora_adv = st.number_input('Tarifa energia fora ponta (opcional)', min_value=0.0, value=0.7, step=0.01, key='adv_tf')
        tarifa_demanda_fora_adv = st.number_input('Tarifa demanda fora ponta (opcional)', min_value=0.0, value=0.0, step=0.01, key='adv_tdf')
else:
    b1, b2 = st.columns(2)
    with b1:
        energia_total_mes_b = st.number_input('Energia total kWh/mês', min_value=0.0, value=5000.0, step=10.0, key='adv_et')
        demanda_estimada_kw = st.number_input('Demanda estimada kW (opcional)', min_value=0.0, value=0.0, step=1.0, key='adv_de')
    with b2:
        tarifa_energia_b_adv = st.number_input('Tarifa energia (opcional)', min_value=0.0, value=0.8, step=0.01, key='adv_tb')

btn1, btn2, btn3 = st.columns(3)
run_fit_adv = btn1.button('1. Gerar curva ajustada')
run_area_adv = btn2.button('2. Definir área disponível')
run_calc_adv = btn3.button('3. Calcular sistema fotovoltaico')

if run_fit_adv:
    selected = profile_map[tipo_cliente_nome]
    peak_labels = classify_peak_hours(list(range(24)), peak_start, peak_end, weekdays_only=weekdays_only_peak)
    if grupo_adv == 'A':
        bill = {
            'energia_ponta_kwh_mes': energia_ponta_mes,
            'energia_fora_ponta_kwh_mes': energia_fora_mes,
            'demanda_ponta_kw': demanda_ponta_kw,
            'demanda_fora_ponta_kw': demanda_fora_kw,
        }
    else:
        bill = {'energia_total_kwh_mes': energia_total_mes_b, 'demanda_estimada_kw': demanda_estimada_kw}

    fit_adv = fit_load_profile_to_bill(
        base_profile_24h=selected['curva_24h_pu'],
        bill_data=bill,
        operation_days=int(dias_func_mes),
        peak_hours=peak_labels,
        group_type=grupo_adv,
    )
    st.session_state['fit_adv'] = fit_adv
    st.session_state['fit_adv_profile'] = selected
    st.success('Curva ajustada gerada com sucesso.')

if st.session_state.get('fit_adv'):
    fit_adv = st.session_state['fit_adv']
    selected = st.session_state['fit_adv_profile']
    st.line_chart(pd.DataFrame({'curva_tipica_pu': selected['curva_24h_pu'], 'curva_ajustada_kw': fit_adv['load_curve_kw']}))
    st.json({k: v for k, v in fit_adv.items() if k != 'load_curve_kw'})

if run_area_adv:
    try:
        if ',' in endereco_mapa and len(endereco_mapa.split(',')) == 2:
            lat_map, lon_map = float(endereco_mapa.split(',')[0]), float(endereco_mapa.split(',')[1])
        else:
            lat_map, lon_map = geocode_address(endereco_mapa)
        st.session_state['adv_map_center'] = (lat_map, lon_map)
    except Exception as e:
        st.error(f'Falha ao geocodificar endereço: {e}')

if st.session_state.get('adv_map_center'):
    map_data = render_area_map(st.session_state['adv_map_center'])
    drawing = map_data.get('last_active_drawing') if isinstance(map_data, dict) else None
    if drawing and drawing.get('geometry', {}).get('type') == 'Polygon':
        coords = drawing['geometry']['coordinates'][0]
        latlon_coords = [(c[1], c[0]) for c in coords[:-1]]
        area_m2 = calculate_polygon_area_m2(latlon_coords)
        st.session_state['available_area_m2'] = area_m2
        st.success(f'Área disponível estimada: {area_m2:.2f} m²')

if st.session_state.get('available_area_m2'):
    st.metric('Área disponível (m²)', f"{st.session_state['available_area_m2']:.2f}")

if run_calc_adv:
    fit_adv = st.session_state.get('fit_adv')
    available_area_m2 = st.session_state.get('available_area_m2')
    if fit_adv is None:
        st.error('Execute o passo 1 para gerar a curva ajustada.')
    elif not available_area_m2:
        st.error('Defina a área disponível no mapa (passo 2).')
    else:
        adjusted = adjusted_profile_from_fit(tipo_cliente_nome, fit_adv)
        target_monthly = adjusted.monthly_energy_kwh
        resultados_df_adv, erro_adv = realizar_dimensionamento_completo(target_monthly, latitude, longitude, int(azimuth), float(tilt))
        if erro_adv:
            st.error(erro_adv)
        else:
            resultados_df_adv, area_meta = aplicar_restricao_area_resultados(
                resultados_df_adv,
                available_area_m2=available_area_m2,
                packing_factor=packing_factor,
            )

            best = resultados_df_adv.iloc[0]
            for warning in area_meta.get('warnings', []):
                st.warning(warning)

            pv_hourly = calcular_geracao_horaria_pvwatts(
                potencia_dc_kwp=float(best['sistema_potencia_total_w']) / 1000.0,
                latitude=latitude,
                longitude=longitude,
                azimuth=int(azimuth),
                tilt=float(tilt),
            ) or [0.0] * 8760
            pv_day = pv_hourly[:24]
            load_day = adjusted.hourly_kw
            peak_labels = classify_peak_hours(list(range(24)), peak_start, peak_end, weekdays_only=weekdays_only_peak)

            sc = analyze_self_consumption(load_day, pv_day)
            ps = analyze_peak_shaving(load_day, pv_day, peak_labels)
            ls = simulate_simplified_load_shifting(
                energy_peak_kwh=fit_adv['energy_peak_kwh_estimated'],
                energy_offpeak_kwh=fit_adv['energy_offpeak_kwh_estimated'],
                percentual_carga_deslocavel_ponta=float(pct_shift) / 100.0,
                shift_window=shift_window,
                tarifa_ponta=tarifa_energia_ponta_adv if grupo_adv == 'A' else None,
                tarifa_fora_ponta=tarifa_energia_fora_adv if grupo_adv == 'A' else None,
            )

            st.subheader('4. Resultado final (novo modo)')
            m1, m2, m3, m4 = st.columns(4)
            m1.metric('Potência recomendada (kWp)', f"{float(best['sistema_potencia_total_w'])/1000.0:.2f}")
            m2.metric('Módulos', f"{int(best['sistema_num_total_paineis'])}")
            m3.metric('Área necessária estimada (m²)', f"{int(best['sistema_num_total_paineis']) * area_meta['module_area_m2']:.2f}")
            m4.metric('Área disponível (m²)', f"{available_area_m2:.2f}")

            st.dataframe(resultados_df_adv.head(10), use_container_width=True)
            st.line_chart(pd.DataFrame({'carga_ajustada_kw': load_day, 'geracao_pv_kw': pv_day, 'carga_liquida_kw': ps['net_load_kw']}))
            st.bar_chart(pd.DataFrame({'autoconsumo_kwh': sc['self_consumption_kwh'], 'injecao_kwh': sc['grid_export_kwh']}))

            st.json({
                'autoconsumo': {
                    'total_load_kwh': sc['total_load_kwh'],
                    'total_pv_generation_kwh': sc['total_pv_generation_kwh'],
                    'total_self_consumption_kwh': sc['total_self_consumption_kwh'],
                    'total_grid_import_kwh': sc['total_grid_import_kwh'],
                    'total_grid_export_kwh': sc['total_grid_export_kwh'],
                    'self_consumption_ratio': sc['self_consumption_ratio'],
                    'self_sufficiency_ratio': sc['self_sufficiency_ratio'],
                },
                'peak_shaving': ps,
                'load_shifting': ls,
                'fit_alerts': fit_adv.get('warnings', []),
            })
