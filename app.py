import streamlit as st
import pandas as pd
from pv_calculator import (
    realizar_dimensionamento_completo, 
    carregar_dados_equipamentos, 
    geocode_location,
    salvar_novo_painel,
    salvar_novo_inversor
)
import inspect

# --- Configurações da Página ---
st.set_page_config(
    page_title="Dimensionamento Fotovoltaico Integrado",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Estilo Personalizado (Preferência do Usuário: Fundo Azul Marinho, Letras Brancas) ---
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

# --- Variáveis Globais ---
FILE_PATH_EQUIPAMENTOS = 'BDFotovoltaica.xlsx'

# --- Inicialização do Session State ---
if 'latitude' not in st.session_state:
    st.session_state['latitude'] = -20.46
if 'longitude' not in st.session_state:
    st.session_state['longitude'] = -54.62
if 'search_lat' not in st.session_state:
    st.session_state['search_lat'] = None
if 'search_lon' not in st.session_state:
    st.session_state['search_lon'] = None

# --- Funções Auxiliares ---
def load_data():
    """Carrega os dados de equipamentos e armazena em cache."""
    try:
        df_paineis, df_inversores = carregar_dados_equipamentos(FILE_PATH_EQUIPAMENTOS)
        return df_paineis, df_inversores
    except Exception as e:
        st.error(f"Erro ao carregar o arquivo de equipamentos ({FILE_PATH_EQUIPAMENTOS}): {e}")
        return pd.DataFrame(), pd.DataFrame()

def apply_coordinates():
    """Aplica as coordenadas encontradas nos inputs principais."""
    if st.session_state['search_lat'] is not None and st.session_state['search_lon'] is not None:
        st.session_state['latitude_input'] = st.session_state['search_lat']
        st.session_state['longitude_input'] = st.session_state['search_lon']
        st.session_state['latitude'] = st.session_state['search_lat']
        st.session_state['longitude'] = st.session_state['search_lon']
        st.session_state['search_lat'] = None
        st.session_state['search_lon'] = None
        st.success(f"Coordenadas aplicadas: Lat={st.session_state['latitude']:.4f}, Lon={st.session_state['longitude']:.4f}")
    else:
        st.warning("Nenhuma coordenada para aplicar. Por favor, realize a busca primeiro.")

def search_coordinates(location_name):
    """Busca as coordenadas e armazena no session state de busca."""
    if location_name:
        with st.spinner(f"Buscando coordenadas para '{location_name}'..."):
            lat_geo, lon_geo = geocode_location(location_name)
            if lat_geo is not None and lon_geo is not None:
                st.session_state['search_lat'] = lat_geo
                st.session_state['search_lon'] = lon_geo
                st.success(f"Localização encontrada: Lat={lat_geo:.4f}, Lon={lon_geo:.4f}. Clique em 'Aplicar Coordenadas' para usar.")
            else:
                st.error("Não foi possível encontrar as coordenadas para a localização informada.")
                st.session_state['search_lat'] = None
                st.session_state['search_lon'] = None
    else:
        st.warning("Por favor, digite o nome da localização.")

# --- Barra Lateral (Sidebar) ---
st.sidebar.title("Configurações do Sistema")

# Informações do Desenvolvedor
st.sidebar.markdown("---")
st.sidebar.markdown("Desenvolvido por **Matheus Vianna**")
st.sidebar.markdown("[matheusvianna.com](https://matheusvianna.com)")
st.sidebar.markdown("---")

# Carregar dados de equipamentos para exibição e seleção
df_paineis, df_inversores = load_data()

with st.sidebar.expander("Dados de Equipamentos Carregados"):
    st.subheader("Painéis Solares")
    if not df_paineis.empty:
        st.dataframe(df_paineis[['modelo', 'potencia_maxima_nominal_pmax', 'tensao_circuito_aberto_voc']].head())
    else:
        st.info("Base de painéis vazia.")
    st.subheader("Inversores")
    if not df_inversores.empty:
        st.dataframe(df_inversores[['modelo', 'maxima_potencia_nominal_ca', 'tensao_maxima_cc']].head())
    else:
        st.info("Base de inversores vazia.")

# --- Formulário de Inserção de Equipamentos ---
with st.sidebar.expander("➕ Inserir Novo Equipamento"):
    st.subheader("Novo Painel Solar")
    with st.form("form_novo_painel", clear_on_submit=True):
        modelo_p = st.text_input("Modelo", key="modelo_p")
        fabricante_p = st.text_input("Fabricante", key="fabricante_p")
        potencia_maxima_nominal_pmax = st.number_input("Potência Máxima Nominal (Pmax) [Wp]", min_value=1.0, step=1.0, format="%.2f", key="pmax_p")
        tensao_circuito_aberto_voc = st.number_input("Tensão de Circuito Aberto (Voc) [V]", min_value=1.0, step=0.1, format="%.2f", key="voc_p")
        corrente_curto_circuito_isc = st.number_input("Corrente de Curto Circuito (Isc) [A]", min_value=1.0, step=0.1, format="%.2f", key="isc_p")
        tensao_operacao_otima_vmp = st.number_input("Tensão de Operação Ótima (Vmp) [V]", min_value=1.0, step=0.1, format="%.2f", key="vmp_p")
        corrente_operacao_otima_imp = st.number_input("Corrente de Operação Ótima (Imp) [A]", min_value=1.0, step=0.1, format="%.2f", key="imp_p")
        eficiencia_modulo = st.number_input("Eficiência do Módulo [%]", min_value=1.0, max_value=100.0, step=0.1, format="%.2f", key="eficiencia_p")
        submitted_painel = st.form_submit_button("Salvar Painel")
        if submitted_painel:
            if modelo_p and potencia_maxima_nominal_pmax > 0 and tensao_circuito_aberto_voc > 0 and corrente_curto_circuito_isc > 0:
                novo_painel_data = {
                    'modelo': modelo_p,
                    'fabricante': fabricante_p,
                    'potencia_maxima_nominal_pmax': potencia_maxima_nominal_pmax,
                    'tensao_operacao_otima_vmp': tensao_operacao_otima_vmp,
                    'corrente_operacao_otima_imp': corrente_operacao_otima_imp,
                    'tensao_circuito_aberto_voc': tensao_circuito_aberto_voc,
                    'corrente_curto_circuito_isc': corrente_curto_circuito_isc,
                    'eficiencia_modulo': eficiencia_modulo
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
        maxima_potencia_nominal_ca = st.number_input("Potência Nominal CA Máxima [W]", min_value=1.0, step=1.0, format="%.2f", key="pot_ca_i")
        tensao_maxima_cc = st.number_input("Tensão Máxima CC [V]", min_value=1.0, step=1.0, format="%.0f", key="vmax_cc_i")
        tensao_start = st.number_input("Tensão de Start [V]", min_value=1.0, step=1.0, format="%.0f", key="vstart_i")
        corrente_maxima_entrada_por_mpp_tracker = st.number_input("Corrente Máxima de Entrada por MPPT [A]", min_value=1.0, step=0.1, format="%.2f", key="imax_mppt_i")
        numero_mpp_trackers = st.number_input("Número de MPPT Trackers", min_value=1, step=1, format="%d", key="num_mppt_i")
        potencia_maxima_fv_maxima = st.number_input("Potência Máxima FV (Máxima) [W]", min_value=1.0, step=1.0, format="%.2f", key="pot_fv_max_i")
        tensao_nominal = st.number_input("Tensão Nominal [V]", min_value=1.0, step=1.0, format="%.0f", key="v_nom_i")
        faixa_tensao_mpp = st.text_input("Faixa de Tensão MPPT (Ex: 60V-550V)", key="faixa_mpp_i")
        corrente_maxima_curto_circuito_por_mpp_tracker = st.number_input("Corrente Máxima Curto Circuito por MPPT [A]", min_value=1.0, step=0.1, format="%.2f", key="isc_mppt_i")
        potencia_maxima_aparente_ca = st.number_input("Potência Máxima Aparente CA [VA]", min_value=1.0, step=1.0, format="%.2f", key="pot_ap_ca_i")
        tensao_nominal_ca = st.number_input("Tensão Nominal CA [V]", min_value=1.0, step=1.0, format="%.0f", key="v_nom_ca_i")
        frequencia_rede_ca = st.text_input("Frequência da Rede CA (Ex: 50Hz/60Hz)", key="freq_ca_i")
        corrente_saida_maxima = st.number_input("Corrente de Saída Máxima [A]", min_value=1.0, step=0.1, format="%.2f", key="i_saida_max_i")
        fator_potencia_ajustavel = st.text_input("Fator de Potência Ajustável (Ex: 0.8i-0.8c)", key="fp_ajustavel_i")
        quantidade_fases_ca = st.number_input("Quantidade de Fases CA", min_value=1, step=1, format="%d", key="fases_ca_i")
        submitted_inversor = st.form_submit_button("Salvar Inversor")
        if submitted_inversor:
            if (modelo_i and maxima_potencia_nominal_ca > 0 and tensao_maxima_cc > 0 and
                tensao_start > 0 and corrente_maxima_entrada_por_mpp_tracker > 0 and numero_mpp_trackers > 0):
                novo_inversor_data = {
                    'modelo': modelo_i,
                    'fabricante': fabricante_i,
                    'potencia_maxima_fv_maxima': potencia_maxima_fv_maxima,
                    'tensao_maxima_cc': tensao_maxima_cc,
                    'tensao_start': tensao_start,
                    'tensao_nominal': tensao_nominal,
                    'faixa_tensao_mpp': faixa_tensao_mpp,
                    'numero_mpp_trackers': numero_mpp_trackers,
                    'corrente_maxima_entrada_por_mpp_tracker': corrente_maxima_entrada_por_mpp_tracker,
                    'corrente_maxima_curto_circuito_por_mpp_tracker': corrente_maxima_curto_circuito_por_mpp_tracker,
                    'maxima_potencia_nominal_ca': maxima_potencia_nominal_ca,
                    'potencia_maxima_aparente_ca': potencia_maxima_aparente_ca,
                    'tensao_nominal_ca': tensao_nominal_ca,
                    'frequencia_rede_ca': frequencia_rede_ca,
                    'corrente_saida_maxima': corrente_saida_maxima,
                    'fator_potencia_ajustavel': fator_potencia_ajustavel,
                    'quantidade_fases_ca': quantidade_fases_ca
                }
                if salvar_novo_inversor(novo_inversor_data):
                    st.success(f"Inversor '{modelo_i}' salvo com sucesso! Recarregue a página para usar.")
                else:
                    st.error("Erro ao salvar o inversor no arquivo Excel.")
            else:
                st.error("Preencha os campos obrigatórios (Modelo, Potência CA, Vmax CC, Vstart, Imax MPPT, Num MPPT).")

# --- Título Principal ---
st.title("☀️ Dimensionamento Fotovoltaico Integrado")
st.markdown("---")

# --- Inputs do Usuário ---
st.header("1. Dados de Geração (PVWatts)")
col1, col2 = st.columns(2)

with col1:
    consumo_medio_mensal = st.number_input(
        "Consumo Médio Mensal (kWh)",
        min_value=1,
        value=300,
        step=10,
        help="Consumo médio mensal de energia elétrica em quilowatt-hora (kWh)."
    )
    location_name = st.text_input(
        "Pesquisar Localização por Nome (Ex: São Paulo, SP)",
        help="Digite o nome da cidade ou endereço para obter as coordenadas."
    )
    if st.button("Buscar Coordenadas"):
        search_coordinates(location_name)

# Botão de Aplicar Coordenadas
if st.session_state['search_lat'] is not None and st.session_state['search_lon'] is not None:
    st.button(
        f"Aplicar Coordenadas Encontradas: Lat={st.session_state['search_lat']:.4f}, Lon={st.session_state['search_lon']:.4f}",
        on_click=apply_coordinates
    )

with col2:
    latitude = st.number_input(
        "Latitude (°)",
        value=st.session_state['latitude'],
        format="%.4f",
        key="latitude_input",
        help="Latitude do local de instalação."
    )
    longitude = st.number_input(
        "Longitude (°)",
        value=st.session_state['longitude'],
        format="%.4f",
        key="longitude_input",
        help="Longitude do local de instalação."
    )
    st.session_state['latitude'] = latitude
    st.session_state['longitude'] = longitude

col3, col4 = st.columns(2)
with col3:
    azimuth = st.selectbox(
        "Azimuth (°)",
        options=[0, 90, 180, 270],
        index=2,
        help="Orientação: 0°=Norte, 90°=Leste, 180°=Sul, 270°=Oeste."
    )
with col4:
    tilt_sugerido = abs(latitude)
    tilt = st.number_input(
        f"Tilt (Inclinação) Sugerido: {tilt_sugerido:.2f}°",
        value=tilt_sugerido,
        format="%.2f",
        help="Inclinação dos painéis em relação ao plano horizontal. Sugerido = latitude."
    )

st.markdown("---")

# --- DEBUG opcional: assinatura da função (ajuda a casar parâmetros) ---
with st.expander("🔧 Diagnóstico (opcional)"):
    try:
        st.write("Assinatura realizar_dimensionamento_completo:", str(inspect.signature(realizar_dimensionamento_completo)))
    except Exception as e:
        st.write("Não foi possível inspecionar a assinatura:", e)

# --- Botão de Cálculo ---
if st.button("Realizar Dimensionamento Completo"):
    if consumo_medio_mensal <= 0:
        st.error("O consumo médio mensal deve ser maior que zero.")
        st.stop()

    if df_paineis.empty or df_inversores.empty:
        st.error("Base de equipamentos vazia. Verifique o arquivo 'BDFotovoltaica.xlsx'.")
        st.stop()

    with st.spinner("Calculando potência de pico e dimensionando o sistema..."):
        try:
            # 🔧 Chamada corrigida: argumentos nomeados, incluindo 'tilt'
            resultados_df, erro = realizar_dimensionamento_completo(
                latitude=latitude,
                longitude=longitude,
                azimuth=azimuth,
                tilt=tilt,
                consumo_medio_mensal=consumo_medio_mensal,
                df_paineis=df_paineis,
                df_inversores=df_inversores,
                # api_key=st.secrets["PVWATTS_API_KEY"],  # descomente se sua função usar
            )
        except TypeError as e:
            st.error(f"Erro de parâmetros na função realizar_dimensionamento_completo: {e}")
            st.stop()
        except Exception as e:
            import sys, traceback
            st.error(f"{type(e).__name__}: {e}")
            st.code(''.join(traceback.format_exception(*sys.exc_info())))
            st.stop()

        if erro:
            st.error(f"Ocorreu um erro durante o dimensionamento: {erro}")
        else:
            st.success("Dimensionamento realizado com sucesso!")

            # --- Resultados ---
            st.header("2. Resultados do Dimensionamento")
            st.subheader("2.1. Resumo de Geração")

            potencia_pico_necessaria_kw = resultados_df['potencia_pico_necessaria_kw'].iloc[0]
            consumo_anual_kwh = resultados_df['consumo_anual_kwh'].iloc[0]
            energia_gerada_anual_kwh = resultados_df['energia_gerada_anual_kwh'].iloc[0]

            col_resumo1, col_resumo2, col_resumo3 = st.columns(3)
            with col_resumo1:
                st.metric("Consumo Anual Alvo (kWh)", f"{consumo_anual_kwh:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            with col_resumo2:
                st.metric("Potência de Pico Necessária (kWp)", f"{potencia_pico_necessaria_kw:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            with col_resumo3:
                st.metric("Energia Anual Estimada (kWh)", f"{energia_gerada_anual_kwh:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

            st.markdown("---")

            # Tabela de Dimensionamento
            st.subheader("2.2. Opções de Dimensionamento (Inversor e Arranjo)")
            df_display = resultados_df[[
                'inversor_modelo', 
                'inversor_fabricante', 
                'inversor_num_unidades', 
                'painel_modelo', 
                'painel_fabricante', 
                'painel_potencia',
                'arranjo_modulos_serie',
                'arranjo_conjuntos_paralelo_por_mppt',
                'inversor_num_mppt',
                'sistema_num_total_paineis',
                'sistema_potencia_total_w'
            ]].copy()

            df_display.columns = [
                'Inversor (Modelo)', 
                'Inversor (Fabricante)', 
                'Inversor (Qtd.)', 
                'Painel (Modelo)', 
                'Painel (Fabricante)', 
                'Painel (Potência Wp)',
                'Módulos em Série (por MPPT)',
                'Conjuntos em Paralelo (por MPPT)',
                'MPPTs (por Inversor)',
                'Total de Painéis',
                'Potência Total do Sistema (Wp)'
            ]

            df_display['Painel (Potência Wp)'] = df_display['Painel (Potência Wp)'].apply(lambda x: f"{x:,.0f}".replace(",", "X").replace(".", ",").replace("X", "."))
            df_display['Potência Total do Sistema (Wp)'] = df_display['Potência Total do Sistema (Wp)'].apply(lambda x: f"{x:,.0f}".replace(",", "X").replace(".", ",").replace("X", "."))
            st.dataframe(df_display, use_container_width=True)

            st.markdown("""
            <div style='color: white; font-size: small;'>
            <strong>Explicação da Tabela:</strong> Cada linha representa uma opção de dimensionamento válida. 
            A coluna 'Potência Total do Sistema (Wp)' indica a potência real instalada, que deve ser próxima da 'Potência de Pico Necessária'.
            O arranjo é detalhado por MPPT (Maximum Power Point Tracker) do inversor.
            </div>
            """, unsafe_allow_html=True)

            st.markdown("---")

            # Detalhes Técnicos (Opcional)
            st.subheader("2.3. Detalhes Técnicos do Melhor Arranjo")
            melhor_arranjo = resultados_df.iloc[0]
            st.markdown(f"""
            <div style='background-color: #004080; padding: 15px; border-radius: 10px;'>
            <strong>Melhor Opção Selecionada:</strong><br><br>
            - <strong>Inversor:</strong> {melhor_arranjo['inversor_modelo']} ({melhor_arranjo['inversor_fabricante']})<br>
            - <strong>Quantidade de Inversores:</strong> {melhor_arranjo['inversor_num_unidades']}<br>
            - <strong>Painel:</strong> {melhor_arranjo['painel_modelo']} ({melhor_arranjo['painel_fabricante']}) - {melhor_arranjo['painel_potencia']:.0f} Wp<br><br>
            <strong>Detalhes do Arranjo (por MPPT):</strong><br>
            - <strong>Módulos em Série:</strong> {melhor_arranjo['arranjo_modulos_serie']}<br>
            - <strong>Conjuntos em Paralelo:</strong> {melhor_arranjo['arranjo_conjuntos_paralelo_por_mppt']}<br>
            - <strong>Potência do Arranjo (por MPPT):</strong> {melhor_arranjo['arranjo_potencia_total_mppt_w']:,.0f} Wp<br><br>
            <strong>Sistema Total:</strong><br>
            - <strong>Potência Total Instalada:</strong> {melhor_arranjo['sistema_potencia_total_w']:,.0f} Wp<br>
            - <strong>Total de Painéis:</strong> {melhor_arranjo['sistema_num_total_paineis']}<br>
            </div>
            """, unsafe_allow_html=True)

            with st.expander("Ver Dados Brutos do Dimensionamento"):
                st.dataframe(resultados_df, use_container_width=True)

# --- Instruções e Explicações (User-Friendly) ---
st.markdown("---")
st.header("Como Funciona o Dimensionamento?")
st.markdown("""
Este aplicativo integra duas etapas cruciais do dimensionamento fotovoltaico:

1.  **Cálculo de Geração (PVWatts):**
    *   Utilizamos a API **PVWatts** do NREL para estimar a produção de energia.
    *   Com base no seu **Consumo Médio Mensal** e nos dados geográficos (**Latitude, Longitude, Azimuth, Tilt**), o sistema calcula a **Potência de Pico Necessária (kWp)** para atender à sua demanda anual.
    *   *O PVWatts simula a irradiação solar e as perdas do sistema para fornecer uma estimativa.*

2.  **Seleção de Inversor e Arranjo:**
    *   Com a Potência de Pico Necessária em mãos, o sistema consulta o **Banco de Dados de Equipamentos** (Painéis e Inversores).
    *   **Seleção do Inversor:** Encontra a combinação de inversores que suporta a potência calculada, considerando margens.
    *   **Cálculo do Arranjo:** Para cada inversor, calcula série/paralelo por MPPT respeitando limites de tensão/corrente.
    *   *O resultado é uma lista de opções de dimensionamento, priorizando a que mais se aproxima da potência alvo.*
""")
