import streamlit as st
import pandas as pd
import inspect
from pv_calculator import (
    realizar_dimensionamento_completo,
    carregar_dados_equipamentos,
    geocode_location,
    salvar_novo_painel,
    salvar_novo_inversor
)

# =========================================================
# CONFIG
# =========================================================
st.set_page_config(
    page_title="Dimensionamento Fotovoltaico Integrado",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================================
# ESTILO
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

FILE_PATH_EQUIPAMENTOS = "BDFotovoltaica.xlsx"

# =========================================================
# SESSION STATE
# =========================================================
if "latitude" not in st.session_state:
    st.session_state["latitude"] = -20.46
if "longitude" not in st.session_state:
    st.session_state["longitude"] = -54.62
if "search_lat" not in st.session_state:
    st.session_state["search_lat"] = None
if "search_lon" not in st.session_state:
    st.session_state["search_lon"] = None

# =========================================================
# HELPERS
# =========================================================
def fmt_num(x, casas=2, vazio="—"):
    try:
        if x is None:
            return vazio
        xv = float(x)
        return f"{xv:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return vazio

def load_data():
    try:
        df_paineis, df_inversores = carregar_dados_equipamentos(FILE_PATH_EQUIPAMENTOS)
        return df_paineis, df_inversores
    except Exception as e:
        st.error(f"Erro ao carregar {FILE_PATH_EQUIPAMENTOS}: {e}")
        return pd.DataFrame(), pd.DataFrame()

def apply_coordinates():
    if st.session_state["search_lat"] is not None and st.session_state["search_lon"] is not None:
        st.session_state["latitude_input"]  = round(st.session_state["search_lat"], 5)
        st.session_state["longitude_input"] = round(st.session_state["search_lon"], 5)
        st.session_state["latitude"]  = st.session_state["latitude_input"]
        st.session_state["longitude"] = st.session_state["longitude_input"]
        st.session_state["search_lat"] = None
        st.session_state["search_lon"] = None
        st.success(
            f"Coordenadas aplicadas: Lat={st.session_state['latitude']:.5f}, "
            f"Lon={st.session_state['longitude']:.5f}"
        )
    else:
        st.warning("Nenhuma coordenada para aplicar. Faça a busca primeiro.")

def search_coordinates(location_name: str):
    if location_name:
        with st.spinner(f"Buscando coordenadas para '{location_name}'..."):
            lat_geo, lon_geo = geocode_location(location_name)
            if lat_geo is not None and lon_geo is not None:
                st.session_state["search_lat"] = lat_geo
                st.session_state["search_lon"] = lon_geo
                st.success(
                    f"Localização encontrada: Lat={lat_geo:.5f}, Lon={lon_geo:.5f}. "
                    "Clique em 'Aplicar Coordenadas' para usar."
                )
            else:
                st.error("Não foi possível encontrar coordenadas para a localização informada.")
                st.session_state["search_lat"] = None
                st.session_state["search_lon"] = None
    else:
        st.warning("Digite o nome da localização.")

def chamar_dimensionamento_seguro(**contexto):
    """
    Chama realizar_dimensionamento_completo mapeando por NOME EXATO da assinatura.
    Normaliza retorno para (df, erro).
    """
    sig = inspect.signature(realizar_dimensionamento_completo)
    params = sig.parameters

    candidatos_base = {
        "latitude": contexto.get("latitude"),
        "longitude": contexto.get("longitude"),
        "azimuth": contexto.get("azimuth"),
        "tilt": contexto.get("tilt"),
        "consumo_medio_mensal": contexto.get("consumo_medio_mensal"),
    }

    args = []
    kwargs = {}
    for nome, p in params.items():
        val = candidatos_base.get(nome)
        if p.kind == inspect.Parameter.POSITIONAL_ONLY:
            if val is None and p.default is inspect._empty:
                raise TypeError(f"Parâmetro obrigatório '{nome}' não fornecido.")
            args.append(val)
        elif p.kind in (inspect.Parameter.POSITIONAL_OR_KEYWORD, inspect.Parameter.KEYWORD_ONLY):
            if val is not None:
                kwargs[nome] = val
            elif p.default is inspect._empty:
                raise TypeError(f"Parâmetro obrigatório '{nome}' não fornecido.")
        else:
            pass

    try:
        resultado = realizar_dimensionamento_completo(*args, **kwargs)
    except TypeError as e:
        raise TypeError(f"Parâmetros incompatíveis ao chamar realizar_dimensionamento_completo "
                        f"com args={args} e kwargs={kwargs}: {e}")

    if isinstance(resultado, tuple) and len(resultado) == 2:
        df, erro = resultado
    else:
        df, erro = resultado, None
    return df, erro

# =========================================================
# SIDEBAR
# =========================================================
st.sidebar.title("Configurações do Sistema")
st.sidebar.markdown("---")
st.sidebar.markdown("Desenvolvido por **Matheus Vianna**")
st.sidebar.markdown("[matheusvianna.com](https://matheusvianna.com)")
st.sidebar.markdown("---")

df_paineis, df_inversores = load_data()

with st.sidebar.expander("Dados de Equipamentos Carregados"):
    st.subheader("Painéis Solares")
    if not df_paineis.empty:
        st.dataframe(df_paineis[["modelo", "potencia_maxima_nominal_pmax", "tensao_circuito_aberto_voc"]].head())
    else:
        st.info("Base de painéis vazia.")
    st.subheader("Inversores")
    if not df_inversores.empty:
        st.dataframe(df_inversores[["modelo", "maxima_potencia_nominal_ca", "tensao_maxima_cc"]].head())
    else:
        st.info("Base de inversores vazia.")

with st.sidebar.expander("➕ Inserir Novo Equipamento"):
    st.subheader("Novo Painel Solar")
    with st.form("form_novo_painel", clear_on_submit=True):
        modelo_p = st.text_input("Modelo", key="modelo_p")
        fabricante_p = st.text_input("Fabricante", key="fabricante_p")
        pmax = st.number_input("Potência Máxima Nominal (Pmax) [Wp]", min_value=1.0, step=1.0, format="%.2f", key="pmax_p")
        voc = st.number_input("Tensão de Circuito Aberto (Voc) [V]", min_value=1.0, step=0.1, format="%.2f", key="voc_p")
        isc = st.number_input("Corrente de Curto Circuito (Isc) [A]", min_value=1.0, step=0.1, format="%.2f", key="isc_p")
        vmp = st.number_input("Tensão de Operação Ótima (Vmp) [V]", min_value=1.0, step=0.1, format="%.2f", key="vmp_p")
        imp = st.number_input("Corrente de Operação Ótima (Imp) [A]", min_value=1.0, step=0.1, format="%.2f", key="imp_p")
        eficiencia = st.number_input("Eficiência do Módulo [%]", min_value=1.0, max_value=100.0, step=0.1, format="%.2f", key="eficiencia_p")
        if st.form_submit_button("Salvar Painel"):
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
                    st.success(f"Painel '{modelo_p}' salvo com sucesso! Recarregue a página para usar.")
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
        imax_mppt = st.number_input("Corrente Máxima de Entrada por MPPT [A]", min_value=1.0, step=0.1, format="%.2f", key="imax_mppt_i")
        num_mppt = st.number_input("Número de MPPT Trackers", min_value=1, step=1, format="%d", key="num_mppt_i")
        pot_fv_max = st.number_input("Potência Máxima FV (Máxima) [W]", min_value=1.0, step=1.0, format="%.2f", key="pot_fv_max_i")
        v_nom = st.number_input("Tensão Nominal [V]", min_value=1.0, step=1.0, format="%.0f", key="v_nom_i")
        faixa_mpp = st.text_input("Faixa de Tensão MPPT (Ex: 60V-550V)", key="faixa_mpp_i")
        isc_mppt = st.number_input("Corrente Máxima Curto Circuito por MPPT [A]", min_value=1.0, step=0.1, format="%.2f", key="isc_mppt_i")
        pot_ap_ca = st.number_input("Potência Máxima Aparente CA [VA]", min_value=1.0, step=1.0, format="%.2f", key="pot_ap_ca_i")
        v_nom_ca = st.number_input("Tensão Nominal CA [V]", min_value=1.0, step=1.0, format="%.0f", key="v_nom_ca_i")
        freq_ca = st.text_input("Frequência da Rede CA (Ex: 50Hz/60Hz)", key="freq_ca_i")
        i_saida_max = st.number_input("Corrente de Saída Máxima [A]", min_value=1.0, step=0.1, format="%.2f", key="i_saida_max_i")
        fp_ajustavel = st.text_input("Fator de Potência Ajustável (Ex: 0.8i-0.8c)", key="fp_ajustavel_i")
        fases_ca = st.number_input("Quantidade de Fases CA", min_value=1, step=1, format="%d", key="fases_ca_i")
        if st.form_submit_button("Salvar Inversor"):
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
                    "corrente_maxima_entrada_por_mpp_tracker": imax_mppt,
                    "corrente_maxima_curto_circuito_por_mpp_tracker": isc_mppt,
                    "maxima_potencia_nominal_ca": pot_ca,
                    "potencia_maxima_aparente_ca": pot_ap_ca,
                    "tensao_nominal_ca": v_nom_ca,
                    "frequencia_rede_ca": freq_ca,
                    "corrente_saida_maxima": i_saida_max,
                    "fator_potencia_ajustavel": fp_ajustavel,
                    "quantidade_fases_ca": fases_ca
                }
                if salvar_novo_inversor(novo_inversor_data):
                    st.success(f"Inversor '{modelo_i}' salvo com sucesso! Recarregue a página para usar.")
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

# Só mostra o botão se realmente temos floats válidos
if (
    st.session_state.get("search_lat") is not None
    and st.session_state.get("search_lon") is not None
):
    try:
        _lat_lbl = float(st.session_state["search_lat"])
        _lon_lbl = float(st.session_state["search_lon"])
        st.button(
            f"Aplicar Coordenadas Encontradas: Lat={_lat_lbl:.5f}, Lon={_lon_lbl:.5f}",
            on_click=apply_coordinates
        )
    except (TypeError, ValueError):
        # Se por algum motivo não forem numéricos, mostra sem formatação
        st.button(
            "Aplicar Coordenadas Encontradas",
            on_click=apply_coordinates
        )

with col2:
    # Garante valores numéricos de fallback para os inputs
    _lat_default = st.session_state.get("latitude")
    try:
        _lat_default = float(_lat_default)
    except (TypeError, ValueError):
        _lat_default = -20.46

    _lon_default = st.session_state.get("longitude")
    try:
        _lon_default = float(_lon_default)
    except (TypeError, ValueError):
        _lon_default = -54.62

    latitude = st.number_input(
        "Latitude (°)",
        value=_lat_default,
        format="%.5f",
        key="latitude_input",
        help="Latitude do local de instalação."
    )
    longitude = st.number_input(
        "Longitude (°)",
        value=_lon_default,
        format="%.5f",
        key="longitude_input",
        help="Longitude do local de instalação."
    )
    st.session_state["latitude"] = latitude
    st.session_state["longitude"] = longitude

col3, col4 = st.columns(2)
with col3:
    azimuth = st.selectbox(
        "Azimuth (°)",
        options=list(range(0, 360, 15)),
        index=12,  # 180 padrão
        help="0°=Norte, 90°=Leste, 180°=Sul, 270°=Oeste."
    )

with col4:
    # Blindagem contra None/strings
    try:
        _lat_for_tilt = float(latitude)
    except (TypeError, ValueError):
        _lat_for_tilt = 0.0
    tilt_sugerido = abs(_lat_for_tilt)

    # Usa label segura (tilt_sugerido é sempre float aqui)
    tilt = st.number_input(
        f"Tilt (Inclinação) Sugerido: {tilt_sugerido:.2f}°",
        value=float(min(max(tilt_sugerido, 0.0), 90.0)),
        min_value=0.0,
        max_value=90.0,
        step=0.5,
        format="%.2f",
        help="Inclinação dos painéis (sugestão = latitude)."
    )

st.markdown("---")

# Diagnóstico opcional
with st.expander("🔧 Diagnóstico (opcional)"):
    try:
        st.write("Assinatura realizar_dimensionamento_completo:", str(inspect.signature(realizar_dimensionamento_completo)))
    except Exception as e:
        st.write("Não foi possível inspecionar a assinatura:", e)

# =========================================================
# BOTÃO
# =========================================================
if st.button("Realizar Dimensionamento Completo"):
    if consumo_medio_mensal <= 0:
        st.error("O consumo médio mensal deve ser maior que zero.")
        st.stop()

    if df_paineis.empty or df_inversores.empty:
        st.warning("Base de equipamentos vazia. Se a função carregar internamente, ignore este aviso.")

    with st.spinner("Calculando potência de pico e dimensionando o sistema..."):
        try:
            resultados_df, erro = chamar_dimensionamento_seguro(
                latitude=latitude,
                longitude=longitude,
                azimuth=azimuth,
                tilt=tilt,
                consumo_medio_mensal=consumo_medio_mensal,
            )
        except Exception as e:
            import sys, traceback
            st.error(f"{type(e).__name__}: {e}")
            st.code("".join(traceback.format_exception(*sys.exc_info())))
            st.stop()

        if erro:
            st.error(f"Ocorreu um erro durante o dimensionamento: {erro}")
            st.stop()

        st.success("Dimensionamento realizado com sucesso!")

        # =========================================================
        # RESULTADOS
        # =========================================================
        st.header("2. Resultados do Dimensionamento")
        st.subheader("2.1. Resumo de Geração")

        potencia_pico_necessaria_kw = resultados_df.get("potencia_pico_necessaria_kw", pd.Series([None])).iloc[0]
        consumo_anual_kwh = resultados_df.get("consumo_anual_kwh", pd.Series([consumo_medio_mensal*12])).iloc[0]
        energia_gerada_anual_kwh = resultados_df.get("energia_gerada_anual_kwh", pd.Series([None])).iloc[0]

        col_resumo1, col_resumo2, col_resumo3 = st.columns(3)
        with col_resumo1:
            st.metric("Consumo Anual Alvo (kWh)", fmt_num(consumo_anual_kwh, 2))
        with col_resumo2:
            st.metric("Potência de Pico Necessária (kWp)", fmt_num(potencia_pico_necessaria_kw, 2))
        with col_resumo3:
            st.metric("Energia Anual Estimada (kWh)", fmt_num(energia_gerada_anual_kwh, 2))

        st.markdown("---")

        # TABELA
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
                lambda x: fmt_num(x, 0)
            )
        if "Potência Total do Sistema (Wp)" in df_display.columns:
            df_display["Potência Total do Sistema (Wp)"] = df_display["Potência Total do Sistema (Wp)"].apply(
                lambda x: fmt_num(x, 0)
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
        melhor_arranjo = resultados_df.iloc[0]

        st.markdown(f"""
        <div style='background-color: #004080; padding: 15px; border-radius: 10px;'>
        <strong>Melhor Opção Selecionada:</strong><br><br>
        - <strong>Inversor:</strong> {melhor_arranjo.get('inversor_modelo', '')} ({melhor_arranjo.get('inversor_fabricante','')})<br>
        - <strong>Quantidade de Inversores:</strong> {melhor_arranjo.get('inversor_num_unidades','')}<br>
        - <strong>Painel:</strong> {melhor_arranjo.get('painel_modelo','')} ({melhor_arranjo.get('painel_fabricante','')}) - {fmt_num(melhor_arranjo.get('painel_potencia',0),0)} Wp<br><br>
        <strong>Detalhes do Arranjo (por MPPT):</strong><br>
        - <strong>Módulos em Série:</strong> {melhor_arranjo.get('arranjo_modulos_serie','')}<br>
        - <strong>Conjuntos em Paralelo:</strong> {melhor_arranjo.get('arranjo_conjuntos_paralelo_por_mppt','')}<br>
        - <strong>Potência do Arranjo (por MPPT):</strong> {fmt_num(melhor_arranjo.get('arranjo_potencia_total_mppt_w',0),0)} Wp<br><br>
        <strong>Sistema Total:</strong><br>
        - <strong>Potência Total Instalada:</strong> {fmt_num(melhor_arranjo.get('sistema_potencia_total_w',0),0)} Wp<br>
        - <strong>Total de Painéis:</strong> {melhor_arranjo.get('sistema_num_total_paineis','')}<br>
        </div>
        """, unsafe_allow_html=True)

        with st.expander("Ver Dados Brutos do Dimensionamento"):
            st.dataframe(resultados_df, use_container_width=True)

# =========================================================
# EXPLICAÇÃO
# =========================================================
st.markdown("---")
st.header("Como Funciona o Dimensionamento?")
st.markdown("""
Este aplicativo integra duas etapas cruciais do dimensionamento fotovoltaico:

1. **Cálculo de Geração (PVWatts):**
   - Usa a API PVWatts (NREL) com `dataset='intl'` para estimar a produção fora dos EUA.
   - A partir de **Consumo Médio Mensal** e **Latitude/Longitude/Azimuth/Tilt**, calcula a **Potência de Pico Necessária (kWp)**.

2. **Seleção de Inversor e Arranjo:**
   - Com a potência alvo, consulta o **banco de dados de Painéis e Inversores**.
   - Define combinações de inversores e arranjos série/paralelo por MPPT respeitando limites elétricos.
   - Retorna opções ordenadas, priorizando a mais próxima da potência alvo.
""")
