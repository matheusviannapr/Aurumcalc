from __future__ import annotations

from datetime import date
from .latex_utils import *


def build_preamble(_: dict) -> str:
    return r"""\documentclass[a4paper,12pt]{article}
\usepackage[utf8]{inputenc}
\usepackage[T1]{fontenc}
\usepackage[brazil]{babel}
\usepackage{geometry,graphicx,booktabs,longtable,array,xcolor,hyperref,siunitx,float,caption,subcaption,amsmath,fancyhdr,pdflscape,multirow}
\geometry{margin=2.3cm}
\pagestyle{fancy}\fancyhf{}\rhead{Memória de Cálculo FV}\lhead{Aurumcalc}\cfoot{\thepage}
\begin{document}
"""


def build_cover(ctx: dict) -> str:
    return f"""\\begin{{titlepage}}\\centering
{{\\Large \\textbf{{MEMÓRIA DE CÁLCULO}}\\\\DIMENSIONAMENTO FOTOVOLTAICO, ANÁLISE HORÁRIA DE CONSUMO E ESTUDO TARIFÁRIO\\par}}
\\vspace{{1cm}}
Cliente: {latex_escape(safe_get(ctx,'project.client_name'))}\\\\
Unidade: {latex_escape(safe_get(ctx,'project.unit_name'))}\\\\
Endereço: {latex_escape(safe_get(ctx,'project.address'))} - {latex_escape(safe_get(ctx,'project.city'))}/{latex_escape(safe_get(ctx,'project.state'))}\\\\
Concessionária: {latex_escape(safe_get(ctx,'utility.company'))}\\\\
Grupo/Modalidade: {latex_escape(safe_get(ctx,'utility.tariff_group'))} / {latex_escape(safe_get(ctx,'utility.tariff_modality'))}\\\\
Responsável técnico: {latex_escape(safe_get(ctx,'project.responsible_engineer'))}\\\\
Empresa: {latex_escape(safe_get(ctx,'project.company'))}\\\\
Data: {latex_escape(safe_get(ctx,'project.report_date', str(date.today())))}\\\\
Versão do software: {latex_escape(safe_get(ctx,'project.software_version'))}
\\end{{titlepage}}"""


def build_table_of_contents() -> str: return "\\tableofcontents\\newpage"

def build_project_identification(ctx: dict) -> str:
    rows = [["Nome do cliente", safe_get(ctx,'project.client_name')],["Unidade consumidora",safe_get(ctx,'project.unit_name')],["Endereço",safe_get(ctx,'project.address')],["Latitude",safe_get(ctx,'project.latitude')],["Longitude",safe_get(ctx,'project.longitude')],["Concessionária",safe_get(ctx,'utility.company')],["Grupo tarifário",safe_get(ctx,'utility.tariff_group')],["Modalidade tarifária",safe_get(ctx,'utility.tariff_modality')],["Tipo de cliente",safe_get(ctx,'load_profile.customer_type')],["Tensão",safe_get(ctx,'utility.voltage_level')],["Período",safe_get(ctx,'utility.billing_days')]]
    return "\\section{Identificação do empreendimento}\n" + build_table(["Campo","Valor"], rows)


def build_objective_section(_: dict) -> str:
    return "\\section{Objetivo da memória de cálculo}Esta memória técnica documenta análise da conta, curva horária, calibração, geração FV, balanço energético, análise ponta/fora ponta, load shifting, peak shaving, área e resultados finais."


def build_assumptions_section(ctx: dict) -> str:
    return "\\section{Normas, referências e premissas}" + build_table(["Item","Valor"], [["Lei 14.300/2022", safe_get(ctx,'assumptions.law_14300','não informado')],["Resoluções ANEEL", safe_get(ctx,'assumptions.aneel','não informado')],["Premissas PVWatts", safe_get(ctx,'pv_generation.source','não informado')],["Perdas totais", format_percent(safe_get(ctx,'pv_system.losses_percent',None))]])


def build_energy_bill_section(ctx: dict) -> str:
    return "\\section{Conta de energia}" + build_table(["Grandeza","Valor"], [["Energia ponta",format_kwh(safe_get(ctx,'bill.energy_peak_kwh',None))],["Energia fora ponta",format_kwh(safe_get(ctx,'bill.energy_offpeak_kwh',None))],["Demanda ponta",format_kw(safe_get(ctx,'bill.demand_peak_kw',None))],["Tarifa energia ponta",format_currency(safe_get(ctx,'bill.tariff_energy_peak',None))]])


def build_load_calibration_section(ctx: dict) -> str:
    tol = safe_get(ctx,'load_calibration.tolerance_percent',None)
    err = safe_get(ctx,'load_calibration.errors.energy_total_percent',None)
    ok = isinstance(tol,(int,float)) and isinstance(err,(int,float)) and abs(err) <= tol
    analysis = "A curva apresenta aderência satisfatória." if ok else "Erro acima da tolerância: recomenda-se medição real ou novo perfil."
    return "\\section{Calibração da curva de demanda}\\[P_{calibrada}(t)=F\\cdot P_{padrao}(t)\\]" + build_table(["Parâmetro","Valor"],[["Fator",safe_get(ctx,'load_calibration.factor')],["Erro energia total",format_percent(err)],["Tolerância",format_percent(tol)]]) + technical_note(analysis)


def build_pv_generation_section(ctx: dict) -> str:
    return "\\section{Geração fotovoltaica}\\[E_{FV}=\\sum_{t=1}^{n}P_{FV}(t)\\]" + build_table(["Indicador","Valor"],[["Geração anual",format_kwh(safe_get(ctx,'pv_generation.annual_kwh',None))],["Geração específica",format_kwh(safe_get(ctx,'pv_generation.specific_yield_kwh_kwp',None))],["Fator de capacidade",format_percent(safe_get(ctx,'pv_generation.capacity_factor_percent',None))]])


def build_hourly_balance_section(ctx: dict) -> str:
    return "\\section{Balanço horário}\\[P_{autoconsumo}(t)=\\min(P_{carga}(t),P_{FV}(t))\\]" + build_table(["Indicador","Valor"],[["Autoconsumo",format_kwh(safe_get(ctx,'hourly_balance.self_consumption_kwh',None))],["Importação",format_kwh(safe_get(ctx,'hourly_balance.grid_import_kwh',None))],["Excedente",format_kwh(safe_get(ctx,'hourly_balance.export_kwh',None))]])


def build_end_document() -> str: return "\\end{document}"
