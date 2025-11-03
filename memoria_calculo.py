# memoria_calculo.py
import json
import math
import datetime
from typing import Optional, Sequence
import pandas as pd

# ---------------------------------------------
# Utilidades
# ---------------------------------------------
MESES_PT = [
    "Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
    "Jul", "Ago", "Set", "Out", "Nov", "Dez"
]

def _fmt_num(x, casas=2):
    try:
        return f"{float(x):,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(x)

def _escape_tex(s: str) -> str:
    if s is None:
        return ""
    # Escapa caracteres comuns do LaTeX
    return (str(s)
            .replace("\\", "\\textbackslash{}")
            .replace("&", "\\&")
            .replace("%", "\\%")
            .replace("$", "\\$")
            .replace("#", "\\#")
            .replace("_", "\\_")
            .replace("{", "\\{")
            .replace("}", "\\}")
            .replace("~", "\\textasciitilde{}")
            .replace("^", "\\textasciicircum{}"))

def _extrair_serie_mensal(resultados_df: pd.DataFrame) -> Optional[Sequence[float]]:
    """
    Busca a coluna 'energia_mensal_kwh_array' (string JSON com 12 valores) na primeira linha.
    Retorna list[float] ou None.
    """
    if resultados_df is None or resultados_df.empty:
        return None
    val = resultados_df.iloc[0].get("energia_mensal_kwh_array", None)
    if val is None:
        return None
    try:
        if isinstance(val, str):
            arr = json.loads(val)
        else:
            arr = val
        if isinstance(arr, (list, tuple)) and len(arr) == 12:
            return [float(x) for x in arr]
        return None
    except Exception:
        return None

def _pick(df, col, default=""):
    try:
        return df.iloc[0].get(col, default)
    except Exception:
        return default

# ---------------------------------------------
# Geração do LaTeX
# ---------------------------------------------
def gerar_memoria_calculo_latex(
    resultados_df: pd.DataFrame,
    caminho_tex: str = "memoria_calculo.tex",
    projeto: str = "Dimensionamento Fotovoltaico",
    cliente: Optional[str] = None,
    local: Optional[str] = None,
    latitude: Optional[float] = None,
    longitude: Optional[float] = None,
    azimuth: Optional[int] = None,
    tilt: Optional[float] = None,
    observacoes: Optional[str] = None,
):
    """
    Gera um arquivo .tex com a memória de cálculo a partir do resultados_df
    retornado por realizar_dimensionamento_completo(...).

    - Inclui: premissas, resumo, arranjo, equipamentos e tabela mensal (sazonalidade).
    - Não depende de pylatex. Usa string LaTeX pura.
    - Não compila o PDF (apenas grava o .tex).
    """
    if resultados_df is None or resultados_df.empty:
        raise ValueError("resultados_df vazio — nada a documentar.")

    # Extrai campos principais
    consumo_anual = _pick(resultados_df, "consumo_anual_kwh", 0)
    potencia_pico_kw = _pick(resultados_df, "potencia_pico_necessaria_kw", 0)
    energia_anual_kwh = _pick(resultados_df, "energia_gerada_anual_kwh", 0)

    inv_modelo = _pick(resultados_df, "inversor_modelo")
    inv_fab = _pick(resultados_df, "inversor_fabricante")
    inv_qtd = _pick(resultados_df, "inversor_num_unidades", 1)
    inv_mppt = _pick(resultados_df, "inversor_num_mppt", 1)

    painel_modelo = _pick(resultados_df, "painel_modelo")
    painel_fab = _pick(resultados_df, "painel_fabricante")
    painel_pot = _pick(resultados_df, "painel_potencia", 0)

    mod_serie = _pick(resultados_df, "arranjo_modulos_serie", 0)
    conj_paralelo = _pick(resultados_df, "arranjo_conjuntos_paralelo_por_mppt", 0)
    pot_mppt_w = _pick(resultados_df, "arranjo_potencia_total_mppt_w", 0)

    pot_total_w = _pick(resultados_df, "sistema_potencia_total_w", 0)
    total_paineis = _pick(resultados_df, "sistema_num_total_paineis", 0)

    energia_mensal = _extrair_serie_mensal(resultados_df)
    consumo_mensal_kwh = None
    if consumo_anual and float(consumo_anual) > 0:
        consumo_mensal_kwh = float(consumo_anual) / 12.0

    # Tabela mensal
    tabela_mensal_linhas = ""
    saz_media = None
    if energia_mensal:
        total_ano_calc = sum(energia_mensal)
        frac = [x / total_ano_calc if total_ano_calc > 0 else 0 for x in energia_mensal]
        saz_media = sum(frac) / 12.0 if frac else None
        for i, kwh in enumerate(energia_mensal):
            tabela_mensal_linhas += (
                f"{MESES_PT[i]} & {_fmt_num(kwh,0)} & "
                f"{_fmt_num(frac[i]*100 if total_ano_calc>0 else 0, 1)}\\% \\\\\n"
            )
    else:
        tabela_mensal_linhas = "\\multicolumn{3}{c}{(Sem dados mensais disponíveis)}\\\\\n"

    # Partes opcionais de cabeçalho
    cliente_tex = f"\\\\ \\textbf{{Cliente:}} {_escape_tex(cliente)}" if cliente else ""
    local_tex = f"\\\\ \\textbf{{Local:}} {_escape_tex(local)}" if local else ""
    coords_tex = ""
    if latitude is not None and longitude is not None:
        coords_tex = f"\\\\ \\textbf{{Coordenadas:}} Lat={latitude:.4f}\\textdegree, Lon={longitude:.4f}\\textdegree"
    data_hoje = datetime.date.today().strftime("%d/%m/%Y")

    # Observações
    obs_tex = ""
    if observacoes:
        obs_tex = (
            "\\section*{Observações}\n"
            f"{_escape_tex(observacoes)}\n\n"
        )

    # Monta LaTeX (documento completo)
    tex = rf"""
\documentclass[a4paper,12pt]{{article}}
\usepackage[utf8]{{inputenc}}
\usepackage[T1]{{fontenc}}
\usepackage[brazil]{{babel}}
\usepackage{{geometry}}
\geometry{{margin=2.5cm}}
\usepackage{{siunitx}}
\sisetup{{output-decimal-marker={{,}}}}
\usepackage{{booktabs}}
\usepackage{{array}}
\usepackage{{graphicx}}
\usepackage{{xcolor}}
\usepackage{{longtable}}
\usepackage{{hyperref}}

\title{{Memória de Cálculo — { _escape_tex(projeto) } }}
\date{{{data_hoje}}}
\author{{}}

\begin{document}
\maketitle

\noindent\textbf{{Projeto:}} { _escape_tex(projeto) }{cliente_tex}{local_tex}{coords_tex}

\section*{{Resumo Executivo}}
Este documento apresenta a memória de cálculo do sistema fotovoltaico proposto, com base nas premissas de consumo e na simulação de geração realizada via PVWatts (NREL). São consolidados o dimensionamento (inversores e arranjos por MPPT), os principais parâmetros técnicos do sistema e a sazonalidade de geração mensal.

\section*{{Premissas e Parâmetros}}
\begin{{itemize}}
  \item Consumo anual alvo: \textbf{{{_fmt_num(consumo_anual_kwh,0)}}} kWh/ano{ " (≈ " + _fmt_num(consumo_mensal_kwh,0) + " kWh/mês)" if consumo_mensal_kwh else "" }.
  \item Potência de pico necessária (PVWatts): \textbf{{{_fmt_num(potencia_pico_kw,2)}}} kWp.
  \item Energia anual estimada do sistema: \textbf{{{_fmt_num(energia_gerada_anual_kwh,0)}}} kWh/ano.
  \item Orientação/Inclinação (entrada do usuário): Azimuth = {azimuth if azimuth is not None else "-"}°; Tilt = { (f"{tilt:.2f}°" if tilt is not None else "-") }.
\end{{itemize}}

\section*{{Dimensionamento — Inversores e Arranjos}}
\subsection*{{Inversores}}
\begin{{itemize}}
  \item Modelo/Fabricante: \textbf{{{_escape_tex(inv_modelo)}}} / \textbf{{{_escape_tex(inv_fab)}}}
  \item Quantidade de inversores: \textbf{{{_escape_tex(inv_qtd)}}}
  \item MPPT por inversor: \textbf{{{_escape_tex(inv_mppt)}}}
\end{{itemize}}

\subsection*{{Módulos Fotovoltaicos}}
\begin{{itemize}}
  \item Modelo/Fabricante: \textbf{{{_escape_tex(painel_modelo)}}} / \textbf{{{_escape_tex(painel_fab)}}}
  \item Potência nominal por módulo: \textbf{{{_fmt_num(painel_pot,0)}}} Wp
  \item Total de módulos no sistema: \textbf{{{_escape_tex(total_paineis)}}}
\end{{itemize}}

\subsection*{{Arranjo por MPPT}}
\begin{{itemize}}
  \item Módulos em série: \textbf{{{_escape_tex(mod_serie)}}}
  \item Conjuntos em paralelo: \textbf{{{_escape_tex(conj_paralelo)}}}
  \item Potência por MPPT: \textbf{{{_fmt_num(pot_mppt_w,0)}}} Wp
\end{{itemize}}

\subsection*{{Potência Global}}
\begin{{itemize}}
  \item Potência total CC estimada: \textbf{{{_fmt_num(pot_total_w,0)}}} Wp
\end{{itemize}}

\section*{{Geração Mensal e Sazonalidade}}
\noindent Abaixo, a distribuição estimada de geração mensal. A soma anual deve aproximar a energia anual estimada do sistema.
\bigskip

\begin{{tabular}}{{lrr}}
\toprule
\textbf{{Mês}} & \textbf{{Geração (kWh)}} & \textbf{{Participação (\%)}} \\
\midrule
{tabela_mensal_linhas}\bottomrule
\end{{tabular}}

\bigskip
\noindent \textit{{Observação:}} Os valores mensais podem variar com clima, sombreamento, degradação de módulos e condições reais de instalação.

{obs_tex}

\section*{{Rastreabilidade}}
Os valores aqui apresentados foram derivados do processo de cálculo do aplicativo (integração PVWatts + base de equipamentos) e correspondem aos resultados retornados pelo dimensionador, considerando as premissas informadas e as limitações técnicas dos inversores e módulos selecionados.

\end{document}
"""

    with open(caminho_tex, "w", encoding="utf-8") as f:
        f.write(tex)

    return caminho_tex
