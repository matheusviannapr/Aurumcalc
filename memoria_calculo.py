# memoria_calculo.py
from datetime import datetime

MESES_PT = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
            "Jul", "Ago", "Set", "Out", "Nov", "Dez"]

def fmt_num(x, casas=2):
    try:
        s = f"{float(x):,.{casas}f}"
        return s.replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(x)

def montar_tabela_mensal_latex(ac_monthly):
    """
    Gera a tabela LaTeX com geração mensal (kWh/mês) e percentual sobre o total.
    """
    if not ac_monthly or len(ac_monthly) != 12:
        return r"\textit{Geração mensal indisponível.}"
    total = sum(ac_monthly)
    linhas = []
    for i, kwh in enumerate(ac_monthly):
        perc = (kwh / total * 100.0) if total > 0 else 0.0
        linhas.append(f"{MESES_PT[i]} & {fmt_num(kwh, 0)} & {fmt_num(perc, 1)}\\% \\\\")
    corpo = "\n".join(linhas)
    tabela = rf"""
\begin{table}[H]
\centering
\caption{{Geração Mensal Estimada (PVWatts)}}
\begin{tabular}{{lrr}}
\toprule
Mês & kWh/mês & Participação \% \\
\midrule
{corpo}
\midrule
\textbf{{Total}} & \textbf{{{fmt_num(total, 0)}}} & \textbf{{100,0\%}} \\
\bottomrule
\end{tabular}
\end{table}
"""
    return tabela

def build_memoria_calculo_latex(payload: dict) -> str:
    """
    payload esperado (chaves principais):
      - projeto: str (ex.: 'PPCI IME')
      - data: str (opcional; default hoje)
      - consumo_mensal_kwh: float
      - consumo_anual_kwh: float
      - latitude, longitude: float
      - azimuth, tilt: float
      - potencia_pico_necessaria_kw: float
      - energia_gerada_anual_kwh: float
      - ac_monthly: lista[12] de floats (kWh/mês)  <-- sazonalidade
      - inversor_modelo, inversor_fabricante, inversor_num_unidades, inversor_num_mppt
      - painel_modelo, painel_fabricante, painel_potencia_w
      - arranjo_modulos_serie, arranjo_conjuntos_paralelo_por_mppt
      - sistema_num_total_paineis, sistema_potencia_total_w
    """
    projeto = payload.get("projeto", "Dimensionamento Fotovoltaico")
    data_hoje = payload.get("data") or datetime.now().strftime("%d/%m/%Y")

    consumo_mensal = payload.get("consumo_mensal_kwh")
    consumo_anual = payload.get("consumo_anual_kwh")
    lat = payload.get("latitude")
    lon = payload.get("longitude")
    az = payload.get("azimuth")
    tilt = payload.get("tilt")

    p_pico_kw = payload.get("potencia_pico_necessaria_kw")
    e_anual = payload.get("energia_gerada_anual_kwh")
    ac_monthly = payload.get("ac_monthly")  # lista de 12

    inv_modelo = payload.get("inversor_modelo")
    inv_fab = payload.get("inversor_fabricante")
    inv_n = payload.get("inversor_num_unidades")
    inv_mppt = payload.get("inversor_num_mppt")

    pnl_modelo = payload.get("painel_modelo")
    pnl_fab = payload.get("painel_fabricante")
    pnl_wp = payload.get("painel_potencia_w")

    serie = payload.get("arranjo_modulos_serie")
    paralelo = payload.get("arranjo_conjuntos_paralelo_por_mppt")
    n_paineis = payload.get("sistema_num_total_paineis")
    pot_total_wp = payload.get("sistema_potencia_total_w")

    tabela_mensal = montar_tabela_mensal_latex(ac_monthly) if ac_monthly else r"\textit{Sem dados mensais}"

    # Documento LaTeX:
    tex = rf"""
\documentclass[12pt,a4paper]{{article}}
\usepackage[utf8]{{inputenc}}
\usepackage[T1]{{fontenc}}
\usepackage[portuguese]{{babel}}
\usepackage{{geometry}}
\geometry{{margin=2.5cm}}
\usepackage{{amsmath, amssymb}}
\usepackage{{booktabs}}
\usepackage{{float}}
\usepackage{{siunitx}}
\sisetup{{locale = FR, output-decimal-marker = {{,}}}}
\usepackage{{hyperref}}
\hypersetup{{colorlinks=true, linkcolor=black, urlcolor=blue}}

\title{{Memória de Cálculo — {projeto}}}
\date{{{data_hoje}}}
\author{{}}

\begin{document}
\maketitle

\section*{{1. Dados de Entrada}}
\begin{itemize}
  \item Consumo médio mensal: \textbf{{{fmt_num(consumo_mensal, 0)}}} kWh/mês
  \item Consumo anual alvo: \textbf{{{fmt_num(consumo_anual, 0)}}} kWh/ano
  \item Latitude / Longitude: \textbf{{{fmt_num(lat, 5)}}} / \textbf{{{fmt_num(lon, 5)}}}
  \item Azimute: \textbf{{{fmt_num(az, 0)}}}° \quad Inclinação (tilt): \textbf{{{fmt_num(tilt, 2)}}}°
\end{itemize}

\section*{{2. Geração PVWatts — Metodologia}}
A produção específica de energia é obtida via simulação PVWatts para um sistema de referência de 1 kW\textsubscript{{p}} no local informado.
Seja $E_{{1kW}}$ (kWh/ano) a energia simulada para 1 kW\textsubscript{{p}}:

\[
E_{{1kW}} = \text{{PVWatts}}(1\,\text{{kW}}, \text{{lat}}, \text{{lon}}, \text{{azimuth}}, \text{{tilt}})
\]

A potência de pico necessária ($P_{{pico}}$) para suprir o consumo anual alvo ($E_{{alvo}}$) é:

\[
P_{{pico}} = \frac{{E_{{alvo}}}}{{E_{{1kW}}}} \quad [\text{{kW}}]
\]

\section*{{3. Resultado de Potência e Geração}}
\begin{itemize}
  \item Potência de pico necessária (kWp): \textbf{{{fmt_num(p_pico_kw, 2)}}}
  \item Energia anual estimada (kWh): \textbf{{{fmt_num(e_anual, 0)}}}
\end{itemize}

{tabela_mensal}

\section*{{4. Seleção de Equipamentos e Arranjo}}
\subsection*{{4.1 Inversor}}
\begin{itemize}
  \item Modelo / Fabricante: \textbf{{{inv_modelo}}} / \textbf{{{inv_fab}}}
  \item Quantidade de inversores: \textbf{{{inv_n}}}
  \item MPPTs por inversor: \textbf{{{inv_mppt}}}
\end{itemize}

\subsection*{{4.2 Módulos FV}}
\begin{itemize}
  \item Modelo / Fabricante: \textbf{{{pnl_modelo}}} / \textbf{{{pnl_fab}}}
  \item Potência do módulo: \textbf{{{fmt_num(pnl_wp, 0)}}} Wp
\end{itemize}

\subsection*{{4.3 Arranjo Elétrico (por MPPT)}}
\begin{itemize}
  \item Módulos em série: \textbf{{{serie}}}
  \item Conjuntos em paralelo: \textbf{{{paralelo}}}
\end{itemize}

\subsection*{{4.4 Sistema Total}}
\begin{itemize}
  \item Total de módulos: \textbf{{{n_paineis}}}
  \item Potência FV agregada: \textbf{{{fmt_num(pot_total_wp, 0)}}} Wp
\end{itemize}

\section*{{5. Observações}}
\begin{itemize}
  \item A sazonalidade mensal apresentada reflete a distribuição típica de irradiação local e perdas consideradas no PVWatts.
  \item Recomenda-se validar o arranjo final considerando limites de tensão CC, corrente por MPPT e normas aplicáveis de instalação.
\end{itemize}

\end{document}
"""
    return tex
