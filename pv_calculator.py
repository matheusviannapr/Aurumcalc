import pandas as pd
import requests
import os
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut, GeocoderServiceError

# --- Configurações e Dados ---
PVWATTS_API_KEY = os.environ.get('PVWATTS_API_KEY', 'hhMhbCRnyXob8EjS7U2AtoTVOjghmjbwDMfDa0Pu')
PVWATTS_URL = 'https://developer.nrel.gov/api/pvwatts/v8.json'
FILE_PATH_EQUIPAMENTOS = 'BDFotovoltaica.xlsx'

# --- Geocode ---
def geocode_location(location_name):
    geolocator = Nominatim(user_agent="pv_dimensioning_app")
    try:
        location = geolocator.geocode(location_name, timeout=10)
        if location:
            # arredonda para evitar cair em “buracos” do dataset
            return round(location.latitude, 5), round(location.longitude, 5)
        return None, None
    except GeocoderTimedOut:
        print("Erro de Geocodificação: Tempo esgotado.")
        return None, None
    except GeocoderServiceError as e:
        print(f"Erro de Geocodificação: {e}")
        return None, None
    except Exception as e:
        print(f"Erro inesperado na Geocodificação: {e}")
        return None, None

# --- PVWatts ---
def fazer_requisicao_pvwatts(params):
    """Chama PVWatts com timeout e surface de erros; força dataset internacional."""
    params['api_key'] = PVWATTS_API_KEY
    params.setdefault('dataset', 'intl')  # essencial para Brasil
    try:
        resp = requests.get(PVWATTS_URL, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        if isinstance(data, dict) and data.get('errors'):
            return {'api_error': True, 'errors': data.get('errors'), 'warnings': data.get('warnings')}
        return data
    except requests.exceptions.RequestException as e:
        return {'api_error': True, 'errors': [str(e)], 'warnings': []}

def _saneia_lat_lon_az_tilt(latitude, longitude, azimuth, tilt):
    lat = max(min(float(latitude), 90.0), -90.0)
    lon = max(min(float(longitude), 180.0), -180.0)
    az  = int(max(min(float(azimuth), 359.0), 0.0))
    tl  = max(min(float(tilt), 90.0), 0.0)
    return round(lat, 4), round(lon, 4), az, tl

def calcular_potencia_pico_necessaria(consumo_medio_mensal, latitude, longitude, azimuth, tilt,
                                      module_type=0, losses=14, array_type=1):
    """Retorna (potencia_pico_kW, json_da_simulacao_1kW_ou_obj_erro)"""
    consumo_anual = consumo_medio_mensal * 12
    lat, lon, az, tl = _saneia_lat_lon_az_tilt(latitude, longitude, azimuth, tilt)

    params = {
        'system_capacity': 1,      # 1 kW para obter produção específica
        'module_type': module_type,
        'losses': losses,
        'array_type': array_type,
        'tilt': tl,
        'azimuth': az,
        'lat': lat,
        'lon': lon,
        'dataset': 'intl',
    }

    data = fazer_requisicao_pvwatts(params)
    if isinstance(data, dict) and data.get('api_error'):
        return None, data

    try:
        annual_energy_per_kw = data['outputs']['ac_annual']
        if annual_energy_per_kw and annual_energy_per_kw > 0:
            potencia_pico_necessaria = consumo_anual / annual_energy_per_kw
            return float(potencia_pico_necessaria), data
    except Exception:
        pass

    return None, {'api_error': True, 'errors': ['Sem ac_annual no retorno PVWatts'], 'warnings': data.get('warnings') if isinstance(data, dict) else []}

def calcular_energia_gerada(potencia_pico, latitude, longitude, azimuth, tilt,
                            module_type=0, losses=14, array_type=1):
    lat, lon, az, tl = _saneia_lat_lon_az_tilt(latitude, longitude, azimuth, tilt)
    params = {
        'system_capacity': float(potencia_pico),
        'module_type': module_type,
        'losses': losses,
        'array_type': array_type,
        'tilt': tl,
        'azimuth': az,
        'lat': lat,
        'lon': lon,
        'dataset': 'intl',
    }
    data = fazer_requisicao_pvwatts(params)
    if isinstance(data, dict) and data.get('api_error'):
        return None
    try:
        return float(data['outputs']['ac_annual'])
    except Exception:
        return None

# --- BD Equipamentos ---
def carregar_dados_equipamentos(file_path):
    try:
        xls = pd.ExcelFile(file_path)
        df_paineis = pd.read_excel(xls, sheet_name='paineis_solares')
        df_inversores = pd.read_excel(xls, sheet_name='Inversores')
        return df_paineis, df_inversores
    except FileNotFoundError:
        print(f"Erro: Arquivo {file_path} não encontrado.")
        return pd.DataFrame(), pd.DataFrame()
    except Exception as e:
        print(f"Erro ao carregar dados do Excel: {e}")
        return pd.DataFrame(), pd.DataFrame()

def salvar_dados_equipamentos(df_paineis, df_inversores, file_path):
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df_paineis.to_excel(writer, sheet_name='paineis_solares', index=False)
            df_inversores.to_excel(writer, sheet_name='Inversores', index=False)
        return True
    except Exception as e:
        print(f"Erro ao salvar dados no Excel: {e}")
        return False

def salvar_novo_painel(novo_painel_data):
    df_paineis, df_inversores = carregar_dados_equipamentos(FILE_PATH_EQUIPAMENTOS)
    novo_painel_df = pd.DataFrame([novo_painel_data])
    colunas_paineis = df_paineis.columns.tolist()
    novo_painel_df = novo_painel_df.reindex(columns=colunas_paineis, fill_value=None)
    df_paineis_atualizado = pd.concat([df_paineis, novo_painel_df], ignore_index=True)
    return salvar_dados_equipamentos(df_paineis_atualizado, df_inversores, FILE_PATH_EQUIPAMENTOS)

def salvar_novo_inversor(novo_inversor_data):
    df_paineis, df_inversores = carregar_dados_equipamentos(FILE_PATH_EQUIPAMENTOS)
    novo_inversor_df = pd.DataFrame([novo_inversor_data])
    colunas_inversores = df_inversores.columns.tolist()
    novo_inversor_df = novo_inversor_df.reindex(columns=colunas_inversores, fill_value=None)
    df_inversores_atualizado = pd.concat([df_inversores, novo_inversor_df], ignore_index=True)
    return salvar_dados_equipamentos(df_paineis, df_inversores_atualizado, FILE_PATH_EQUIPAMENTOS)

# --- Dimensionamento (inversor/arranjo) ---
def selecionar_inversores(potencia_pico_w, df_inversores):
    potencia_inversor_necessaria = potencia_pico_w / 1.2
    combinacoes_inversores = []

    for _, inversor in df_inversores.iterrows():
        potencia_nominal_ca = inversor.get('maxima_potencia_nominal_ca')
        if pd.isna(potencia_nominal_ca) or float(potencia_nominal_ca) <= 0:
            continue

        potencia_nominal_ca = float(potencia_nominal_ca)
        num_inversores_min = int(potencia_inversor_necessaria // potencia_nominal_ca)
        if potencia_inversor_necessaria % potencia_nominal_ca > 0:
            num_inversores_min += 1

        for num_inversores in range(num_inversores_min, max(num_inversores_min,1) + 2):
            potencia_combinacao = num_inversores * potencia_nominal_ca
            potencia_pico_sistema = potencia_pico_w
            potencia_maxima_cc_suportada = potencia_combinacao * 1.2

            if potencia_pico_sistema * 0.95 <= potencia_maxima_cc_suportada <= potencia_pico_sistema * 1.5:
                combinacoes_inversores.append({
                    'modelo': inversor.get('modelo'),
                    'fabricante': inversor.get('fabricante'),
                    'maxima_potencia_nominal_ca': potencia_nominal_ca,
                    'tensao_maxima_cc': inversor.get('tensao_maxima_cc'),
                    'tensao_start': inversor.get('tensao_start'),
                    'corrente_maxima_entrada_por_mpp_tracker': inversor.get('corrente_maxima_entrada_por_mpp_tracker'),
                    'numero_mpp_trackers': inversor.get('numero_mpp_trackers'),
                    'num_inversores': num_inversores,
                    'potencia_combinacao_ca': potencia_combinacao,
                    'potencia_maxima_cc_suportada': potencia_maxima_cc_suportada,
                    'diferenca_potencia_pico': abs(potencia_maxima_cc_suportada - potencia_pico_sistema)
                })

    df_combinacoes = pd.DataFrame(combinacoes_inversores)
    if df_combinacoes.empty:
        return df_combinacoes
    idx = df_combinacoes.groupby('modelo')['diferenca_potencia_pico'].idxmin()
    return df_combinacoes.loc[idx].sort_values(by='diferenca_potencia_pico')

def calcular_arranjos_possiveis(df_paineis, tensao_maxima_cc, tensao_start):
    arranjos = []
    tmax = float(tensao_maxima_cc or 0)
    tstart = float(tensao_start or 0)

    for _, modulo in df_paineis.iterrows():
        potencia_modulo = float(modulo.get('potencia_maxima_nominal_pmax', 0) or 0)
        tensao_voc = float(modulo.get('tensao_circuito_aberto_voc', 0) or 0)
        corrente_isc = float(modulo.get('corrente_curto_circuito_isc', 0) or 0)
        if tensao_voc <= 0 or corrente_isc <= 0 or potencia_modulo <= 0:
            continue

        max_modulos_serie = int(tmax // tensao_voc)
        for num_modulos_serie in range(1, max_modulos_serie + 1):
            tensao_total = num_modulos_serie * tensao_voc
            if tstart <= tensao_total <= tmax:
                arranjos.append({
                    'modelo_painel': modulo.get('modelo'),
                    'fabricante_painel': modulo.get('fabricante'),
                    'potencia_modulo': potencia_modulo,
                    'num_modulos_serie': num_modulos_serie,
                    'tensao_total_voc': tensao_total,
                    'corrente_isc': corrente_isc
                })
    return pd.DataFrame(arranjos)

def selecionar_arranjo_paineis(potencia_pico_inversor_w, df_paineis, tensao_maxima_cc, tensao_start, corrente_maxima_mppt, num_mppt):
    potencia_necessaria_mppt = float(potencia_pico_inversor_w) / float(num_mppt)
    potencia_minima = potencia_necessaria_mppt * 0.9
    potencia_maxima = potencia_necessaria_mppt * 1.1

    arranjos_possiveis = calcular_arranjos_possiveis(df_paineis, tensao_maxima_cc, tensao_start)
    if arranjos_possiveis.empty:
        return None

    arranjos_validos = []
    corr_max = float(corrente_maxima_mppt or 0)

    for _, arr in arranjos_possiveis.iterrows():
        pot_serie = int(arr['num_modulos_serie']) * float(arr['potencia_modulo'])
        if pot_serie <= 0:
            continue
        max_paralelo_corrente = int(corr_max // float(arr['corrente_isc'])) if arr['corrente_isc'] > 0 else 0
        max_paralelo_pot = int(potencia_maxima // pot_serie) if pot_serie > 0 else 0
        max_paralelo = max(0, min(max_paralelo_corrente, max_paralelo_pot))

        for npar in range(1, max_paralelo + 1):
            pot_total = pot_serie * npar
            corrente_total = float(arr['corrente_isc']) * npar
            if potencia_minima <= pot_total <= potencia_maxima and corrente_total <= corr_max:
                arranjos_validos.append({
                    'modelo_painel': arr['modelo_painel'],
                    'fabricante_painel': arr['fabricante_painel'],
                    'potencia_modulo': arr['potencia_modulo'],
                    'num_modulos_serie': int(arr['num_modulos_serie']),
                    'num_conjuntos_paralelo': int(npar),
                    'potencia_total_arranjo_mppt': float(pot_total),
                    'corrente_total_isc': float(corrente_total),
                    'diferenca_potencia_mppt': abs(pot_total - potencia_necessaria_mppt),
                })

    if not arranjos_validos:
        return None
    dfv = pd.DataFrame(arranjos_validos)
    return dfv.loc[dfv['diferenca_potencia_mppt'].idxmin()].to_dict()

def dimensionar_sistema(potencia_pico_w, df_paineis, df_inversores):
    melhores = selecionar_inversores(potencia_pico_w, df_inversores)
    if melhores.empty:
        return None, "Nenhuma combinação de inversores válida encontrada."

    resultados = []
    for _, comb in melhores.iterrows():
        pot_por_inv = potencia_pico_w / float(comb['num_inversores'])
        arr_ideal = selecionar_arranjo_paineis(
            pot_por_inv,
            df_paineis,
            comb['tensao_maxima_cc'],
            comb['tensao_start'],
            comb['corrente_maxima_entrada_por_mpp_tracker'],
            comb['numero_mpp_trackers']
        )
        if arr_ideal:
            pot_total_mppt = float(arr_ideal['potencia_total_arranjo_mppt'])
            num_mppt = int(comb['numero_mpp_trackers'])
            q_inv = int(comb['num_inversores'])
            pot_total = pot_total_mppt * num_mppt * q_inv

            n_serie = int(arr_ideal['num_modulos_serie'])
            n_par = int(arr_ideal['num_conjuntos_paralelo'])
            total_paineis = n_serie * n_par * num_mppt * q_inv

            resultados.append({
                'inversor_modelo': comb['modelo'],
                'inversor_fabricante': comb['fabricante'],
                'inversor_potencia_ca_nominal': float(comb['maxima_potencia_nominal_ca']),
                'inversor_num_unidades': q_inv,
                'inversor_num_mppt': num_mppt,
                'painel_modelo': arr_ideal['modelo_painel'],
                'painel_fabricante': arr_ideal['fabricante_painel'],
                'painel_potencia': float(arr_ideal['potencia_modulo']),
                'arranjo_modulos_serie': n_serie,
                'arranjo_conjuntos_paralelo_por_mppt': n_par,
                'arranjo_potencia_total_mppt_w': pot_total_mppt,
                'sistema_potencia_total_w': pot_total,
                'sistema_num_total_paineis': total_paineis,
                'sistema_potencia_pico_alvo_w': float(potencia_pico_w),
            })

    if not resultados:
        return None, "Nenhum arranjo de painéis válido encontrado."
    return pd.DataFrame(resultados), None

# --- Função principal ---
def realizar_dimensionamento_completo(consumo_medio_mensal, latitude, longitude, azimuth, tilt):
    potencia_pico_kw, dados_pvwatts = calcular_potencia_pico_necessaria(
        consumo_medio_mensal, latitude, longitude, azimuth, tilt
    )
    if not potencia_pico_kw:
        msg = "Falha ao calcular a potência de pico necessária via PVWatts."
        if isinstance(dados_pvwatts, dict) and dados_pvwatts.get('api_error'):
            errs = "; ".join(dados_pvwatts.get('errors') or [])
            warns = "; ".join(dados_pvwatts.get('warnings') or [])
            if errs:
                msg += f" [errors: {errs}]"
            if warns:
                msg += f" [warnings: {warns}]"
        return None, msg

    potencia_pico_w = float(potencia_pico_kw) * 1000.0

    df_paineis, df_inversores = carregar_dados_equipamentos(FILE_PATH_EQUIPAMENTOS)
    if df_paineis.empty or df_inversores.empty:
        return None, "Falha ao carregar dados de painéis ou inversores. Verifique o arquivo de equipamentos."

    df_dim, erro = dimensionar_sistema(potencia_pico_w, df_paineis, df_inversores)
    if erro:
        return None, erro

    energia_gerada_anual = calcular_energia_gerada(potencia_pico_kw, latitude, longitude, azimuth, tilt)
    # Pode ser None se a PVWatts falhar aqui — o app formata isso com segurança

    df_dim['consumo_anual_kwh'] = float(consumo_medio_mensal) * 12.0
    df_dim['potencia_pico_necessaria_kw'] = float(potencia_pico_kw)
    df_dim['energia_gerada_anual_kwh'] = energia_gerada_anual  # pode ser None
    return df_dim, None

if __name__ == '__main__':
    consumo = 1000
    lat = -23.55
    lon = -46.64
    azimuth = 0
    tilt = abs(lat)
    res, err = realizar_dimensionamento_completo(consumo, lat, lon, azimuth, tilt)
    if err:
        print("Erro:", err)
    else:
        print(res.head())
