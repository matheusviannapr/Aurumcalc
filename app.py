import streamlit as st
import pandas as pd
from pv_calculator import realizar_dimensionamento_completo, carregar_dados_equipamentos

# --- Configurações da Página ---
st.set_page_config(
    page_title="Dimensionamento Fotovoltaico Integrado",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Estilo Personalizado (Preferência do Usuário: Fundo Azul Marinho, Letras Brancas) ---
# O Streamlit não permite alterar o fundo principal para azul marinho diretamente via st.set_page_config
# sem usar CSS. Vamos adicionar um bloco de CSS para aplicar o estilo.
custom_css = """
<style>
    /* Estilo para o fundo principal e texto */
    .stApp {
        background-color: #001f3f; /* Azul Marinho Escuro */
        color: white;
    }
    /* Estilo para o texto em geral */
    p, h1, h2, h3, h4, h5, h6, .stMarkdown {
        color: white;
    }
    /* Estilo para a barra lateral */
    .css-1d391kg, .css-1lcbmhc { /* Classes para a barra lateral */
        background-color: #003366; /* Um tom de azul mais claro para a sidebar */
        color: white;
    }
    /* Estilo para os inputs (para garantir que o texto seja legível) */
    .stTextInput > div > div > input, .stNumberInput > div > div > input {
        background-color: #004080;
        color: white;
    }
    /* Estilo para os labels dos inputs */
    .stForm label, .stTextInput label, .stNumberInput label, .stSelectbox label {
        color: white;
    }
    /* Estilo para os botões */
    .stButton>button {
        background-color: #007bff;
        color: white;
        border-radius: 5px;
    }
    .stButton>button:hover {
        background-color: #0056b3;
        color: white;
    }
    /* Estilo para as tabelas/dataframes */
    .dataframe {
        color: black !important; /* Cor do texto da tabela para contraste */
        background-color: white !important; /* Fundo da tabela */
    }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# --- Variáveis Globais ---
FILE_PATH_EQUIPAMENTOS = 'BDFotovoltaica.xlsx'

# --- Barra Lateral (Sidebar) ---
st.sidebar.title("Configurações do Sistema")

# Informações do Desenvolvedor
st.sidebar.markdown("---")
st.sidebar.markdown("Desenvolvido por **Matheus Vianna**")
st.sidebar.markdown("[matheusvianna.com](https://matheusvianna.com)")
st.sidebar.markdown("---")

# Carregar dados de equipamentos para exibição e seleção
try:
    df_paineis, df_inversores = carregar_dados_equipamentos(FILE_PATH_EQUIPAMENTOS)
    
    # Exibir dados carregados na sidebar (opcional, para debug ou referência)
    with st.sidebar.expander("Dados de Equipamentos Carregados"):
        st.subheader("Painéis Solares")
        st.dataframe(df_paineis[['modelo', 'potencia_maxima_nominal_pmax', 'tensao_circuito_aberto_voc']].head())
        st.subheader("Inversores")
        st.dataframe(df_inversores[['modelo', 'maxima_potencia_nominal_ca', 'tensao_maxima_cc']].head())

except Exception as e:
    st.error(f"Erro ao carregar o arquivo de equipamentos ({FILE_PATH_EQUIPAMENTOS}): {e}")
    st.stop()

# --- Título Principal ---
st.title("☀️ Dimensionamento Fotovoltaico Integrado")
st.markdown("---")

# --- Inputs do Usuário ---

# Seção de Geração (PVWatts)
st.header("1. Dados de Geração (PVWatts)")
col1, col2, col3 = st.columns(3)

with col1:
    consumo_medio_mensal = st.number_input(
        "Consumo Médio Mensal (kWh)",
        min_value=1,
        value=300,
        step=10,
        help="Consumo médio mensal de energia elétrica em quilowatt-hora (kWh)."
    )
    latitude = st.number_input(
        "Latitude (°)",
        value=-20.46,
        format="%.2f",
        help="Latitude do local de instalação."
    )

with col2:
    longitude = st.number_input(
        "Longitude (°)",
        value=-54.62,
        format="%.2f",
        help="Longitude do local de instalação."
    )
    azimuth = st.selectbox(
        "Azimuth (°)",
        options=[0, 90, 180, 270],
        index=2, # 180 (Norte)
        help="Orientação dos painéis em relação ao Norte (0°=Norte, 90°=Leste, 180°=Sul, 270°=Oeste)."
    )

with col3:
    # Sugestão de Tilt ideal (igual à latitude)
    tilt_sugerido = abs(latitude)
    tilt = st.number_input(
        f"Tilt (Inclinação) Sugerido: {tilt_sugerido:.2f}°",
        value=tilt_sugerido,
        format="%.2f",
        help="Inclinação dos painéis em relação ao plano horizontal. O valor sugerido é a latitude do local."
    )

st.markdown("---")

# --- Botão de Cálculo ---
if st.button("Realizar Dimensionamento Completo"):
    
    # Validação básica
    if consumo_medio_mensal <= 0:
        st.error("O consumo médio mensal deve ser maior que zero.")
        st.stop()
        
    with st.spinner("Calculando potência de pico e dimensionando o sistema..."):
        
        # Executar a função principal de integração
        resultados_df, erro = realizar_dimensionamento_completo(
            consumo_medio_mensal, 
            latitude, 
            longitude, 
            azimuth, 
            tilt, 
            FILE_PATH_EQUIPAMENTOS
        )
        
        if erro:
            st.error(f"Ocorreu um erro durante o dimensionamento: {erro}")
        else:
            st.success("Dimensionamento realizado com sucesso!")
            
            # --- Resultados ---
            st.header("2. Resultados do Dimensionamento")
            
            # Informações de Geração
            st.subheader("2.1. Resumo de Geração")
            
            potencia_pico_necessaria_kw = resultados_df['potencia_pico_necessaria_kw'].iloc[0]
            consumo_anual_kwh = resultados_df['consumo_anual_kwh'].iloc[0]
            energia_gerada_anual_kwh = resultados_df['energia_gerada_anual_kwh'].iloc[0]
            
            col_resumo1, col_resumo2, col_resumo3 = st.columns(3)
            
            with col_resumo1:
                st.metric(
                    label="Consumo Anual Alvo (kWh)", 
                    value=f"{consumo_anual_kwh:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                )
            with col_resumo2:
                st.metric(
                    label="Potência de Pico Necessária (kWp)", 
                    value=f"{potencia_pico_necessaria_kw:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                )
            with col_resumo3:
                st.metric(
                    label="Energia Anual Estimada (kWh)", 
                    value=f"{energia_gerada_anual_kwh:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                )
                
            st.markdown("---")
            
            # Tabela de Dimensionamento
            st.subheader("2.2. Opções de Dimensionamento (Inversor e Arranjo)")
            
            # Selecionar e renomear colunas para melhor visualização
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
            
            # Formatar colunas numéricas
            df_display['Painel (Potência Wp)'] = df_display['Painel (Potência Wp)'].apply(lambda x: f"{x:,.0f}".replace(",", "X").replace(".", ",").replace("X", "."))
            df_display['Potência Total do Sistema (Wp)'] = df_display['Potência Total do Sistema (Wp)'].apply(lambda x: f"{x:,.0f}".replace(",", "X").replace(".", ",").replace("X", "."))
            
            st.dataframe(df_display, use_container_width=True)
            
            st.markdown("""
            <div style='color: white; font-size: small;'>
            **Explicação da Tabela:** Cada linha representa uma opção de dimensionamento válida. 
            A coluna 'Potência Total do Sistema (Wp)' indica a potência real instalada, que deve ser próxima da 'Potência de Pico Necessária'.
            O arranjo é detalhado por MPPT (Maximum Power Point Tracker) do inversor.
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("---")
            
            # Detalhes Técnicos (Opcional)
            st.subheader("2.3. Detalhes Técnicos do Melhor Arranjo")
            
            # O melhor arranjo é o primeiro da lista (o que tem a menor diferença de potência)
            melhor_arranjo = resultados_df.iloc[0]
            
            st.markdown(f"""
            <div style='background-color: #004080; padding: 15px; border-radius: 10px;'>
            
            **Melhor Opção Selecionada:**
            
            - **Inversor:** {melhor_arranjo['inversor_modelo']} ({melhor_arranjo['inversor_fabricante']})
            - **Quantidade de Inversores:** {melhor_arranjo['inversor_num_unidades']}
            - **Painel:** {melhor_arranjo['painel_modelo']} ({melhor_arranjo['painel_fabricante']}) - {melhor_arranjo['painel_potencia']:.0f} Wp
            
            **Detalhes do Arranjo (por MPPT):**
            
            - **Módulos em Série:** {melhor_arranjo['arranjo_modulos_serie']}
            - **Conjuntos em Paralelo:** {melhor_arranjo['arranjo_conjuntos_paralelo_por_mppt']}
            - **Potência do Arranjo (por MPPT):** {melhor_arranjo['arranjo_potencia_total_mppt_w']:,.0f} Wp
            
            **Sistema Total:**
            
            - **Potência Total Instalada:** {melhor_arranjo['sistema_potencia_total_w']:,.0f} Wp
            - **Total de Painéis:** {melhor_arranjo['sistema_num_total_paineis']}
            
            </div>
            """, unsafe_allow_html=True)
            
            # Exibir o DataFrame completo para análise detalhada (opcional)
            with st.expander("Ver Dados Brutos do Dimensionamento"):
                st.dataframe(resultados_df, use_container_width=True)
                
# --- Instruções e Explicações (Melhoria de User-Friendly) ---
st.markdown("---")
st.header("Como Funciona o Dimensionamento?")
st.markdown("""
Este aplicativo integra duas etapas cruciais do dimensionamento fotovoltaico:

1.  **Cálculo de Geração (PVWatts):**
    *   Utilizamos a API **PVWatts** do NREL (National Renewable Energy Laboratory) para estimar a produção de energia.
    *   Com base no seu **Consumo Médio Mensal** e nos dados geográficos (**Latitude, Longitude, Azimuth, Tilt**), o sistema calcula a **Potência de Pico Necessária (kWp)** para atender à sua demanda anual.
    *   *O PVWatts simula a irradiação solar e as perdas do sistema para fornecer uma estimativa precisa.*

2.  **Seleção de Inversor e Arranjo:**
    *   Com a Potência de Pico Necessária em mãos, o sistema consulta o **Banco de Dados de Equipamentos** (Painéis e Inversores).
    *   **Seleção do Inversor:** Encontra a combinação ideal de inversores que suporta a potência calculada, considerando uma margem de segurança (overload).
    *   **Cálculo do Arranjo:** Para cada inversor, calcula o número ideal de painéis em **série** e em **paralelo** (o arranjo) que respeita as restrições elétricas do inversor (Tensão Máxima CC, Tensão de Start e Corrente Máxima do MPPT).
    *   *O resultado é uma lista de opções de dimensionamento, priorizando a que mais se aproxima da potência alvo.*
""")
