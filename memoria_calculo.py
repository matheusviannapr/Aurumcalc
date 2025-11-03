# memoria_calculo.py
import json
from datetime import datetime
from typing import Any, List, Tuple, Union

import pandas as pd


def _as_dataframe(resultados: Union[pd.DataFrame, dict, List[dict]]) -> pd.DataFrame:
    """Normaliza 'resultados' para DataFrame."""
    if isinstance(resultados, pd.DataFrame):
        return resultados.copy()

    if isinstance(resultados, dict):
        # dict de colunas -> listas OU um único registro
        # Tenta inferir: se valores não são listas, vira uma linha
        if not resultados:
            return pd.DataFrame()
        any_val = next(iter(resultados.values()))
        if isinstance(any_val, list):
            return pd.DataFrame(resultados)
        else:
            return pd.DataFrame([resultados])

    if isinstance(resultados, list):
        # lista de dicts
        return pd.DataFrame(resultados)

    # fallback
    return pd.DataFrame()


def _parse_monthly_array(val: Any) -> List[float]:
    """
    Normaliza a energia mensal (12 valores) para lista[float].
    Aceita: lista, tuple, string JSON, ou None -> retorna lista vazia.
    """
    if val is None:
        return []

    # Já é lista/tupla de números?
    if isinstance(val, (list, tuple)):
        try:
            return [float(x) for x in val]
        except Exception:
            return []

    # Tenta decodificar JSON em string
    if isinstance(val, str):
        try:
            arr = json.loads(val)
            if isinstance(arr, (list, tuple)):
                return [float(x) for x in arr]
        except Exception:
            return []

    # Caso não reconheça o formato
    return []


def _coalesce_row_value(row: pd.Series, keys: List[str], default: float = 0.0) -> float:
    """Lê, em ordem, a primeira coluna existente e converte para float."""
    for k in keys:
        if k in row and pd.notna(row[k]):
            try:
                return float(row[k])
            except Exception:
                pass
    return default


def gerar_memoria_calculo_latex(
    resultados_df: Union[pd.DataFrame, dict, List[dict]],
    caminho_tex: str = "memoria_calculo.tex",
    projeto: str = "Dimensionamento FV",
    cliente: str = "Cliente",
    local: str = "",
    latitude: float = None,
    longitude: float = None,
    azimuth: float = None,
    tilt: float = None,
    observacoes: str = "",
) -> str:
    """
    Gera um arquivo LaTeX com memória de cálculo do dimensionamento.
    - resultados_df pode ser DataFrame, dict, ou list[dict].
    - Espera (idealmente) as colunas:
        potencia_pico_necessaria_kw, consumo_anual_kwh, energia_gerada_anual_kwh,
        energia_mensal_kwh_array (lista ou JSON), além dos dados do melhor arranjo.
    Retorna o caminho do arquivo .tex gerado.
    """
    df = _as_dataframe(resultados_df)

    if df.shape[0] == 0:
        raise ValueError("resultados_df está vazio (sem linhas).")

    # Considera a primeira linha como “selecionada”
    row = df.iloc[0]

    # Campos principais (com chaves alternativas para maior tolerância)
    potencia_kw = _coalesce_row_value(row, ["potencia_pico_necessaria_kw", "potencia_pico_kw"], 0.0)
    consumo_anual = _coalesce_row_value(row, ["consumo_anual_kwh"], 0.0)
    energia_anual = _coalesce_row_value(row, ["energia_gerada_anual_kwh", "energia_anual_kwh"], 0.0)

    energia_mensal = []
    # Tenta pegar um array mensal (12 valores)
    for key in ["energia_mensal_kwh_array", "ac_monthly_kwh", "ac_monthly", "energia_mensal_kwh"]:
        if key in row:
            energia_mensal = _parse_monthly_array(row[key])
            break

    # Metadados do arranjo (se existirem)
    inversor_modelo = row.get("inversor_modelo", "")
    inversor_fabricante = row.get("inversor_fabricante", "")
    inversor_num_unid = row.get("inversor_num_unidades", "")
    inversor_num_mppt = row.get("inversor_num_mppt", "")

    painel_modelo = row.get("painel_modelo", "")
    painel_fabricante = row.get("painel_fabricante", "")
    painel_potencia = row.get("painel_potencia", "")

    mod_serie = row.get("arranjo_modulos_serie", "")
    conj_paralelo = row.get("arranjo_conjuntos_paralelo_por_mppt", "")
    pot_mppt_w = row.get("arranjo_potencia_total_mppt_w", "")
    pot_total_w = row.get("sistema_potencia_total_w", "")
    num_paineis = row.get("sistema_num_total_paineis", "")

    # Formatações “seguras”
    def fmt(v, nd=2, milhar=True):
        try:
            if v is None or (isinstance(v, float) and pd.isna(v)):
                return "-"
            if isinstance(v, (int, float)):
                s = f"{v:,.{nd}f}" if milhar else f"{v:.{nd}f}"
                # pt-BR
                s = s.replace(",", "X").replace(".", ",").replace("X", ".")
                return s
            # string
            return str(v)
        except Exception:
            return str(v)

    # Tabela mensal (se houver)
    meses = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
    tem_mensal = len(energia_mensal) == 12

    hoje = datetime.now().strftime("%d/%m/%Y")

    # Monta o LaTeX
    # Obs: documento simples para evitar dependências além do básico
    conteudo = r"""\documentclass[a4paper,12pt]{article}
\usepackage[utf8]{inputenc}
\usepackage[brazil]{babel}
\usepackage{geometry}
\usepackage{array,booktabs,longtable}
\usepackage{siunitx}
\usepackage{graphicx}
\usepackage{float}
\usepackage{setspace}
\geometry{margin=2.5cm}
\sisetup{
  output-decimal-marker = {,}, 
  group-separator = {.}
}
\begin{document}
\begin{center}
\Large \textbf{Memória de Cálculo — Sistema Fotovoltaico}\\[4pt]
\large %(PROJETO)s
\end{center}

\noindent
\textbf{Cliente:} %(CLIENTE)s \\
\textbf{Local:} %(LOCAL)s \\
\textbf{Data:} %(DATA)s

\vspace{1em}
\section*{Premissas de Geração}
\begin{itemize}
  \item Latitude: %(LAT)s\textdegree
  \item Longitude: %(LON)s\textdegree
  \item Azimute: %(AZI)s\textdegree
  \item Inclinação (Tilt): %(TILT)s\textdegree
\end{itemize}

\section*{Resultados Principais}
\begin{tabular}{@{}ll@{}}
\toprule
Consumo anual alvo (kWh)                  & %(CONSUMO_ANUAL)s \\
Potência de pico necessária (kWp)         & %(POT_KWP)s \\
Energia anual estimada (kWh)              & %(ENERGIA_ANUAL)s \\
\bottomrule
\end{tabular}

\vspace{1em}
\section*{Configuração Selecionada (Resumo)}
\begin{tabular}{@{}ll@{}}
\toprule
Inversor (Modelo/Fabricante) & %(INV_MODEL)s / %(INV_FAB)s \\
Quantidade de Inversores     & %(INV_QTD)s \\
MPPTs (por inversor)         & %(INV_MPPT)s \\
Painel (Modelo/Fabricante)   & %(PAN_MODEL)s / %(PAN_FAB)s \\
Potência do Painel (Wp)      & %(PAN_POT)s \\
Módulos em Série (por MPPT)  & %(MOD_SERIE)s \\
Conjuntos em Paralelo (MPPT) & %(CONJ_PARA)s \\
Potência por MPPT (Wp)       & %(POT_MPPT_W)s \\
Potência Total Instalada (Wp)& %(POT_TOTAL_W)s \\
Total de Painéis             & %(TOT_PAIN)s \\
\bottomrule
\end{tabular}
"""  # noqa: E501

    # Bloco mensal opcional
    if tem_mensal:
        conteudo += r"""

\vspace{1em}
\section*{Sazonalidade — Energia Mensal Estimada (kWh)}
\begin{tabular}{@{}l*{12}{r}@{}}
\toprule
Mês & Jan & Fev & Mar & Abr & Mai & Jun & Jul & Ago & Set & Out & Nov & Dez \\
\midrule
Energia (kWh) & %(M1)s & %(M2)s & %(M3)s & %(M4)s & %(M5)s & %(M6)s & %(M7)s & %(M8)s & %(M9)s & %(M10)s & %(M11)s & %(M12)s \\
\bottomrule
\end{tabular}
"""
    conteudo += r"""

\vspace{1em}
\section*{Observações}
%(OBS)s

\end{document}
"""

    # Monta o dicionário de substituição
    sub = {
        "PROJETO": projeto,
        "CLIENTE": cliente,
        "LOCAL": local,
        "DATA": hoje,
        "LAT": "-" if latitude is None else fmt(latitude, nd=4, milhar=False),
        "LON": "-" if longitude is None else fmt(longitude, nd=4, milhar=False),
        "AZI": "-" if azimuth is None else fmt(azimuth, nd=0, milhar=False),
        "TILT": "-" if tilt is None else fmt(tilt, nd=1, milhar=False),
        "CONSUMO_ANUAL": fmt(consumo_anual, nd=2),
        "POT_KWP": fmt(potencia_kw, nd=2),
        "ENERGIA_ANUAL": fmt(energia_anual, nd=2),
        "INV_MODEL": str(inversor_modelo or "-"),
        "INV_FAB": str(inversor_fabricante or "-"),
        "INV_QTD": str(inversor_num_unid or "-"),
        "INV_MPPT": str(inversor_num_mppt or "-"),
        "PAN_MODEL": str(painel_modelo or "-"),
        "PAN_FAB": str(painel_fabricante or "-"),
        "PAN_POT": fmt(painel_potencia, nd=0),
        "MOD_SERIE": str(mod_serie or "-"),
        "CONJ_PARA": str(conj_paralelo or "-"),
        "POT_MPPT_W": fmt(pot_mppt_w, nd=0),
        "POT_TOTAL_W": fmt(pot_total_w, nd=0),
        "TOT_PAIN": str(num_paineis or "-"),
        "OBS": observacoes or "-",
    }

    if tem_mensal:
        for i in range(12):
            sub[f"M{i+1}"] = fmt(energia_mensal[i], nd=0)

    # Escreve o arquivo
    with open(caminho_tex, "w", encoding="utf-8") as f:
        f.write(conteudo % sub)

    return caminho_tex
