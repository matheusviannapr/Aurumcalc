# pv_calculator.py
import os
import json
import math
import pandas as pd
import requests

# geocoding (opcional: se não houver geopy instalado, o geocode falha com None,None)
try:
    from geopy.geocoders import Nominatim
    from geopy.exc import GeocoderTimedOut, GeocoderServiceError
    HAS_GEOPY = True
except Exception:
    HAS_GEOPY = False

# ---------------------------------------------------------
# Configurações
# ---------------------------------------------------------
PVWATTS_API_KEY = os.environ.get("PVWATTS_API_KEY", None)
PVWATTS_URL = "https://developer.nrel.gov/api/pvwatts/v8.json"

FILE_PATH_EQUIPAMENTOS_DEFAULT = "BDFotovoltaica.xlsx"
SHEET_PAIN_EIS = "paineis_solares"
SHEET_INV = "Inversores"

# Defaults técnicos (podem ser refinados na sua lógica)
DEFAULT_MODULE_TYPE = 0   # 0=standard, 1=premium, 2=thin film
DEFAULT_LOSSES = 14       # % (total system losses)
DEFAULT_ARRAY_TYPE = 1    # 1=fixed roof


# ---------------------------------------------------------
# Utilidades
# ---------------------------------------------------------
def _is_number(x):
    try:
        float(x)
        return True
    except Exception:
        return False


# ---------------------------------------------------------
# Geocodificação
# ---------------------------------------------------------
def geocode_location(location_name: str):
    """
    Converte nome de local em (latitude, longitude) via Nominatim (OpenStreetMap).
    Retorna (None, None) se falhar ou se geopy não estiver disponível.
    """
    if not HAS_GEOPY or not location_name:
        return None, None
    try:
        geolocator = Nominatim(user_agent="pv_dimensioning_app")
        location = geolocator.geocode(location_name, timeout=10)
        if location:
            return float(location.latitude), float(location.longitude)
        return None, None
    except (GeocoderTimedOut, GeocoderServiceError):
        return None, None
    except Exception:
        return None, None


# ---------------------------------------------------------
# PVWatts
# ---------------------------------------------------------
def fazer_requisicao_pvwatts(params: dict):
    """
    Chama a API PVWatts com checagem de erros.
    Retorna dict (JSON) ou None.
    """
    api_key = os.environ.get("PVWATTS_API_KEY", PVWATTS_API_KEY)
    if not api_key:
        # Sem chave, não tem como consultar
        return None

    all_params = dict(params)
    all_params["api_key"] = api_key

    try:
        resp = requests.get(PVWATTS_URL, params=all_params, timeout=20)
        resp.raise_for_status()
        data = resp.json()
        # PVWatts retorna 'errors'/'warnings' dentro de 'outputs' às vezes
        if isinstance(data, dict) and "errors" in data.get("outputs", {}):
            # Se houver erro explícito nos outputs, trate como falha
            errs = data["outputs"]["errors"]
            if isinstance(errs, list) and len(errs) > 0:
                return None
        return data
    except requests.exceptions.RequestException:
        return None
    except ValueError:
        return None


def _extrair_ac_monthly(data: dict):
    """
    Tenta extrair ac_monthly do retorno do PVWatts.
    Retorna list[float] ou None.
    """
    try:
        arr = data["outputs"]["ac_monthly"]
        if isinstance(arr, list) and len(arr) == 12:
            # garante floats
            return [float(x) for x in arr]
        return None
    except Exception:
        return None


def calcular_potencia_pico_necessaria(
    consumo_medio_mensal_kwh: float,
    latitude: float,
    longitude: float,
    azimuth: int,
    tilt: float,
    module_type: int = DEFAULT_MODULE_TYPE,
    losses: float = DEFAULT_LOSSES,
    array_type: int = DEFAULT_ARRAY_TYPE,
):
    """
    Calcula kWp necessários usando PVWatts via simulação de 1 kW.

    Retorna:
      (potencia_pico_necessaria_kw: float | None,
       dados_pvwatts_1kw: dict | None,
       ac_monthly_1kw: list[float] | None)
    """
    if consumo_medio_mensal_kwh is None or consumo_medio_mensal_kwh <= 0:
        return None, None, None
    if not (_is_number(latitude) and _is_number(longitude)):
        return None, None, None

    consumo_anual_kwh = float(consumo_medio_mensal_kwh) * 12.0

    params = {
        "system_capacity": 1.0,  # 1 kW para obter kWh/ano por kWp
        "module_type": int(module_type),
        "losses": float(losses),
        "array_type": int(array_type),
        "tilt": float(tilt),
        "azimuth": int(azimuth),
        "lat": float(latitude),
        "lon": float(longitude),
    }

    data = fazer_requisicao_pvwatts(params)
    if not data:
        return None, None, None

    ac_annual = None
    try:
        ac_annual = float(data["outputs"]["ac_annual"])
    except Exception:
        ac_annual = None

    if not ac_annual or ac_annual <= 0:
        return None, None, None

    potencia_kw = consumo_anual_kwh / ac_annual  # (kWh/ano) / (kWh/ano por 1 kW)
    ac_monthly = _extrair_ac_monthly(data)  # pode ser None

    return float(potencia_kw), data, ac_monthly


def calcular_energia_gerada(
    potencia_pico_kw: float,
    latitude: float,
    longitude: float,
    azimuth: int,
    tilt: float,
    module_type: int = DEFAULT_MODULE_TYPE,
    losses: float = DEFAULT_LOSSES,
    array_type: int = DEFAULT_ARRAY_TYPE,
):
    """
    Calcula energia anual (kWh) para uma potência de pico (kW).
    Retorna (energia_anual_kwh: float | None, ac_monthly: list[float] | None)
    """
    if not potencia_pico_kw or potencia_pico_kw <= 0:
        return None, None

    params = {
        "system_capacity": float(potencia_pico_kw),
        "module_type": int(module_type),
        "losses": float(losses),
        "array_type": int(array_type),
        "tilt": float(tilt),
        "azimuth": int(azimuth),
        "lat": float(latitude),
        "lon": float(longitude),
    }
    data = fazer_requisicao_pvwatts(params)
    if not data:
        return None, None

    try:
        ac_annual = float(data["outputs"]["ac_annual"])
    except Exception:
        ac_annual = None

    ac_monthly = _extrair_ac_monthly(data)

    return ac_annual, ac_monthly


# ---------------------------------------------------------
# Equipamentos: carregar/salvar (Excel)
# ---------------------------------------------------------
def carregar_dados_equipamentos(file_path: str = FILE_PATH_EQUIPAMENTOS_DEFAULT):
    """
    Retorna (df_paineis, df_inversores). Se falhar, retorna (DataFrame vazio, DataFrame vazio).
    """
    try:
        xls = pd.ExcelFile(file_path)
        df_paineis = pd.read_excel(xls, sheet_name=SHEET_PAIN_EIS)
        df_inversores = pd.read_excel(xls, sheet_name=SHEET_INV)
        return df_paineis, df_inversores
    except FileNotFoundError:
        return pd.DataFrame(), pd.DataFrame()
    except Exception:
        return pd.DataFrame(), pd.DataFrame()


def _salvar_excel_duas_abas(df_paineis: pd.DataFrame, df_inversores: pd.DataFrame, file_path: str):
    try:
        # exige openpyxl
        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            df_paineis.to_excel(writer, sheet_name=SHEET_PAIN_EIS, index=False)
            df_inversores.to_excel(writer, sheet_name=SHEET_INV, index=False)
        return True
    except Exception:
        return False


def salvar_novo_painel(novo_painel_data: dict, file_path: str = FILE_PATH_EQUIPAMENTOS_DEFAULT):
    df_paineis, df_inversores = carregar_dados_equipamentos(file_path)
    if df_paineis.empty:
        # se a base ainda não existe, cria colunas a partir das chaves
        df_paineis = pd.DataFrame(columns=list(novo_painel_data.keys()))
    novo_df = pd.DataFrame([novo_painel_data]).reindex(columns=df_paineis.columns, fill_value=None)
    df_paineis = pd.concat([df_paineis, novo_df], ignore_index=True)
    if df_inversores.empty:
        df_inversores = pd.DataFrame(columns=[
            "modelo", "fabricante", "maxima_potencia_nominal_ca", "tensao_maxima_cc", "tensao_start",
            "corrente_maxima_entrada_por_mpp_tracker", "numero_mpp_trackers",
            "corrente_maxima_curto_circuito_por_mpp_tracker", "potencia_maxima_fv_maxima",
            "tensao_nominal", "faixa_tensao_mpp", "potencia_maxima_aparente_ca",
            "tensao_nominal_ca", "frequencia_rede_ca", "corrente_saida_maxima",
            "fator_potencia_ajustavel", "quantidade_fases_ca"
        ])
    return _salvar_excel_duas_abas(df_paineis, df_inversores, file_path)


def salvar_novo_inversor(novo_inversor_data: dict, file_path: str = FILE_PATH_EQUIPAMENTOS_DEFAULT):
    df_paineis, df_inversores = carregar_dados_equipamentos(file_path)
    if df_inversores.empty:
        df_inversores = pd.DataFrame(columns=list(novo_inversor_data.keys()))
    novo_df = pd.DataFrame([novo_inversor_data]).reindex(columns=df_inversores.columns, fill_value=None)
    df_inversores = pd.concat([df_inversores, novo_df], ignore_index=True)
    if df_paineis.empty:
        df_paineis = pd.DataFrame(columns=[
            "modelo", "fabricante", "potencia_maxima_nominal_pmax", "tensao_operacao_otima_vmp",
            "corrente_operacao_otima_imp", "tensao_circuito_aberto_voc", "corrente_curto_circuito_isc",
            "eficiencia_modulo"
        ])
    return _salvar_excel_duas_abas(df_paineis, df_inversores, file_path)


# ---------------------------------------------------------
# Dimensionamento: inversores e arranjos
# ---------------------------------------------------------
def selecionar_inversores(potencia_pico_w: float, df_inversores: pd.DataFrame) -> pd.DataFrame:
    """
    Seleciona combinações plausíveis de inversores para atender a potência de pico CC.
    Heurística: potencia_inversor_necessaria = potencia_pico_w / 1.2 (supõe ~20% margem CC->CA).
    Retorna DataFrame com melhores combinações por modelo, ordenadas por diferença.
    """
    if potencia_pico_w is None or potencia_pico_w <= 0 or df_inversores is None or df_inversores.empty:
        return pd.DataFrame()

    potencia_inversor_necessaria = float(potencia_pico_w) / 1.2
    lst = []

    for _, inv in df_inversores.iterrows():
        pot_ca = inv.get("maxima_potencia_nominal_ca", None)
        if pot_ca is None or not _is_number(pot_ca) or float(pot_ca) <= 0:
            continue

        pot_ca = float(pot_ca)

        # número mínimo de inversores para atender a potência necessária
        n_min = max(1, math.ceil(potencia_inversor_necessaria / pot_ca))
        for num in range(n_min, n_min + 3):  # tenta 3 passos acima
            pot_comb_ca = num * pot_ca
            pot_max_cc_suportada = pot_comb_ca * 1.2  # supõe relação aproximada CC/CA

            diferenca = abs(pot_max_cc_suportada - potencia_pico_w)

            lst.append({
                "modelo": inv.get("modelo"),
                "fabricante": inv.get("fabricante"),
                "maxima_potencia_nominal_ca": pot_ca,
                "tensao_maxima_cc": inv.get("tensao_maxima_cc"),
                "tensao_start": inv.get("tensao_start"),
                "corrente_maxima_entrada_por_mpp_tracker": inv.get("corrente_maxima_entrada_por_mpp_tracker"),
                "numero_mpp_trackers": inv.get("numero_mpp_trackers"),
                "num_inversores": num,
                "potencia_combinacao_ca": pot_comb_ca,
                "potencia_maxima_cc_suportada": pot_max_cc_suportada,
                "diferenca_potencia_pico": diferenca
            })

    df = pd.DataFrame(lst)
    if df.empty:
        return df

    # Melhor combinação por modelo (menor diferença)
    try:
        idx = df.groupby("modelo")["diferenca_potencia_pico"].idxmin()
        df = df.loc[idx].sort_values(by="diferenca_potencia_pico").reset_index(drop=True)
    except Exception:
        df = df.sort_values(by="diferenca_potencia_pico").reset_index(drop=True)
    return df


def calcular_arranjos_possiveis(df_paineis: pd.DataFrame, tensao_max_cc: float, tensao_start: float) -> pd.DataFrame:
    """
    Gera possibilidades de módulos em série respeitando janela de tensão do inversor.
    """
    if df_paineis is None or df_paineis.empty or not _is_number(tensao_max_cc) or not _is_number(tensao_start):
        return pd.DataFrame()

    tensao_max_cc = float(tensao_max_cc)
    tensao_start = float(tensao_start)
    lst = []

    for _, m in df_paineis.iterrows():
        pmod = m.get("potencia_maxima_nominal_pmax")
        voc = m.get("tensao_circuito_aberto_voc")
        isc = m.get("corrente_curto_circuito_isc")

        if not (_is_number(voc) and float(voc) > 0 and _is_number(isc) and float(isc) > 0 and _is_number(pmod) and float(pmod) > 0):
            continue

        voc = float(voc)
        isc = float(isc)
        pmod = float(pmod)

        max_n_serie = max(1, int(tensao_max_cc // voc))
        for n in range(1, max_n_serie + 1):
            v_total = n * voc
            if tensao_start <= v_total <= tensao_max_cc:
                lst.append({
                    "modelo_painel": m.get("modelo"),
                    "fabricante_painel": m.get("fabricante"),
                    "potencia_modulo": pmod,
                    "num_modulos_serie": n,
                    "tensao_total_voc": v_total,
                    "corrente_isc": isc
                })
    return pd.DataFrame(lst)


def selecionar_arranjo_paineis(
    potencia_por_inversor_w: float,
    df_paineis: pd.DataFrame,
    tensao_max_cc: float,
    tensao_start: float,
    corrente_max_mppt: float,
    num_mppt: int
):
    """
    Seleciona arranjo (série/paralelo) para UM inversor, repartindo potência alvo por MPPT.
    Heurística: alvo por MPPT = (potencia_por_inversor_w / num_mppt) * [0.9, 1.1]
    """
    if not (potencia_por_inversor_w and potencia_por_inversor_w > 0 and _is_number(num_mppt) and int(num_mppt) > 0):
        return None

    num_mppt = int(num_mppt)
    alvo_mppt_w = float(potencia_por_inversor_w) / num_mppt
    pmin = alvo_mppt_w * 0.90
    pmax = alvo_mppt_w * 1.10

    poss = calcular_arranjos_possiveis(df_paineis, tensao_max_cc, tensao_start)
    if poss.empty:
        return None

    candidatos = []
    for _, a in poss.iterrows():
        p_serie = a["num_modulos_serie"] * a["potencia_modulo"]
        if p_serie <= 0:
            continue

        # Corrente limita o número de paralelos
        if not (_is_number(corrente_max_mppt) and float(corrente_max_mppt) > 0 and _is_number(a["corrente_isc"]) and float(a["corrente_isc"]) > 0):
            continue

        i_max = float(corrente_max_mppt)
        isc = float(a["corrente_isc"])
        max_paralelos_por_corrente = max(1, int(i_max // isc))

        # Potência também limita
        max_paralelos_por_pot = max(1, int(pmax // p_serie))
        max_paralelos = max(1, min(max_paralelos_por_corrente, max_paralelos_por_pot))

        best_local = None
        best_diff = None
        for k in range(1, max_paralelos + 1):
            p_tot = p_serie * k
            i_tot = isc * k
            if pmin <= p_tot <= pmax and i_tot <= i_max:
                diff = abs(p_tot - alvo_mppt_w)
                if best_diff is None or diff < best_diff:
                    best_diff = diff
                    best_local = {
                        "modelo_painel": a["modelo_painel"],
                        "fabricante_painel": a["fabricante_painel"],
                        "potencia_modulo": a["potencia_modulo"],
                        "num_modulos_serie": int(a["num_modulos_serie"]),
                        "num_conjuntos_paralelo": int(k),
                        "potencia_total_arranjo_mppt": float(p_tot),
                        "corrente_total_isc": float(i_tot),
                        "diferenca_potencia_mppt": float(diff),
                    }
        if best_local:
            candidatos.append(best_local)

    if not candidatos:
        return None

    df = pd.DataFrame(candidatos)
    return df.loc[df["diferenca_potencia_mppt"].idxmin()].to_dict()


def dimensionar_sistema(potencia_pico_w: float, df_paineis: pd.DataFrame, df_inversores: pd.DataFrame):
    """
    Integra seleção de inversor + arranjo MPPT, retornando DataFrame consolidado ou (None, erro).
    """
    comb = selecionar_inversores(potencia_pico_w, df_inversores)
    if comb is None or comb.empty:
        return None, "Nenhuma combinação de inversores válida encontrada."

    resultados = []
    for _, c in comb.iterrows():
        # potencia de pico total é compartilhada pelos inversores
        pot_por_inversor_w = float(potencia_pico_w) / float(c["num_inversores"])

        arr = selecionar_arranjo_paineis(
            pot_por_inversor_w,
            df_paineis,
            c.get("tensao_maxima_cc"),
            c.get("tensao_start"),
            c.get("corrente_maxima_entrada_por_mpp_tracker"),
            c.get("numero_mpp_trackers"),
        )
        if not arr:
            # se não encontrou arranjo para esse inversor, tenta próximo
            continue

        num_mppt = int(c.get("numero_mpp_trackers", 1))
        num_inversores = int(c.get("num_inversores", 1))

        pot_total_mppt = arr["potencia_total_arranjo_mppt"]
        pot_total_sistema = float(pot_total_mppt) * num_mppt * num_inversores

        num_mods_serie = int(arr["num_modulos_serie"])
        num_paralelo = int(arr["num_conjuntos_paralelo"])
        total_paineis = num_mods_serie * num_paralelo * num_mppt * num_inversores

        resultados.append({
            "inversor_modelo": c.get("modelo"),
            "inversor_fabricante": c.get("fabricante"),
            "inversor_potencia_ca_nominal": c.get("maxima_potencia_nominal_ca"),
            "inversor_num_unidades": num_inversores,
            "inversor_num_mppt": num_mppt,

            "painel_modelo": arr["modelo_painel"],
            "painel_fabricante": arr["fabricante_painel"],
            "painel_potencia": arr["potencia_modulo"],

            "arranjo_modulos_serie": num_mods_serie,
            "arranjo_conjuntos_paralelo_por_mppt": num_paralelo,
            "arranjo_potencia_total_mppt_w": float(pot_total_mppt),

            "sistema_potencia_total_w": float(pot_total_sistema),
            "sistema_num_total_paineis": int(total_paineis),

            # placeholders (preenchidos na fase final)
            "potencia_pico_necessaria_kw": None,
            "consumo_anual_kwh": None,
            "energia_gerada_anual_kwh": None,

            # payload opcional do PVWatts mensal (preenchemos depois, na função principal)
            "energia_mensal_kwh_array": None,
        })

    if not resultados:
        return None, "Nenhum arranjo de painéis válido encontrado para as combinações de inversores."

    return pd.DataFrame(resultados), None


# ---------------------------------------------------------
# Função Principal de Integração
# ---------------------------------------------------------
def realizar_dimensionamento_completo(
    consumo_medio_mensal: float,
    latitude: float,
    longitude: float,
    azimuth: int,
    tilt: float,
    arquivo_equipamentos: str = FILE_PATH_EQUIPAMENTOS_DEFAULT
):
    """
    1) Calcula kWp necessários pelo PVWatts (simulação 1 kW).
    2) Carrega bases (painéis/inversores).
    3) Dimensiona inversor + arranjo.
    4) Calcula energia anual do sistema dimensionado (kWh) e tenta obter série mensal (kWh/mês).
    5) Retorna (DataFrame, erro|None).
    """
    # 1) kWp necessários
    kWp, data_1kW, ac_monthly_1kw = calcular_potencia_pico_necessaria(
        consumo_medio_mensal, latitude, longitude, azimuth, tilt,
        module_type=DEFAULT_MODULE_TYPE,
        losses=DEFAULT_LOSSES,
        array_type=DEFAULT_ARRAY_TYPE
    )
    if not kWp:
        return None, "Falha ao calcular a potência de pico necessária via PVWatts. Verifique a chave da API, coordenadas ou a conectividade."

    # 2) Carregar bases
    df_paineis, df_inversores = carregar_dados_equipamentos(arquivo_equipamentos)
    if df_paineis.empty or df_inversores.empty:
        return None, "Falha ao carregar dados de painéis ou inversores. Verifique o arquivo de equipamentos."

    # 3) Dimensionar
    potencia_pico_w = float(kWp) * 1000.0
    df_dim, erro = dimensionar_sistema(potencia_pico_w, df_paineis, df_inversores)
    if erro:
        return None, erro

    if df_dim is None or df_dim.empty:
        return None, "Nenhum resultado de dimensionamento foi obtido."

    # 4) Energia anual do sistema dimensionado
    energia_anual_kwh, ac_monthly_dim = calcular_energia_gerada(
        kWp, latitude, longitude, azimuth, tilt,
        module_type=DEFAULT_MODULE_TYPE,
        losses=DEFAULT_LOSSES,
        array_type=DEFAULT_ARRAY_TYPE
    )

    # 5) Completar colunas
    df_dim["potencia_pico_necessaria_kw"] = float(kWp)
    df_dim["consumo_anual_kwh"] = float(consumo_medio_mensal) * 12.0
    df_dim["energia_gerada_anual_kwh"] = energia_anual_kwh if energia_anual_kwh is not None else None

    # tenta abastecer energia_mensal_kwh_array:
    # prioridade: a simulação com kWp real (ac_monthly_dim); fallback: escala ac_monthly_1kw * kWp
    energia_mensal = None
    if ac_monthly_dim and isinstance(ac_monthly_dim, list) and len(ac_monthly_dim) == 12:
        energia_mensal = [float(x) for x in ac_monthly_dim]
    elif ac_monthly_1kw and isinstance(ac_monthly_1kw, list) and len(ac_monthly_1kw) == 12:
        energia_mensal = [float(x) * float(kWp) for x in ac_monthly_1kw]

    # guarda a série mensal SOMENTE na primeira linha (mas você pode replicar se preferir)
    if energia_mensal:
        try:
            df_dim.loc[df_dim.index[0], "energia_mensal_kwh_array"] = json.dumps(energia_mensal)
        except Exception:
            df_dim.loc[df_dim.index[0], "energia_mensal_kwh_array"] = str(energia_mensal)
    else:
        df_dim.loc[df_dim.index[0], "energia_mensal_kwh_array"] = None

    return df_dim, None


# ---------------------------------------------------------
# Execução de teste simples
# ---------------------------------------------------------
if __name__ == "__main__":
    # Teste rápido (ajuste coordenadas ao seu caso)
    consumo = 1000
    lat = -23.5505
    lon = -46.6333
    az = 0
    tilt = abs(lat)

    print(">> Teste PVWatts 1 kW...")
    kWp, data1, acm1 = calcular_potencia_pico_necessaria(consumo, lat, lon, az, tilt)
    print("kWp:", kWp, "ac_monthly_1kW len:", 0 if acm1 is None else len(acm1))

    print(">> Teste dimensionamento geral...")
    df, erro = realizar_dimensionamento_completo(consumo, lat, lon, az, tilt)
    if erro:
        print("ERRO:", erro)
    else:
        print(df.head())
        print("OK.")
