import pandas as pd
import requests
import os
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut, GeocoderServiceError

# --- Configurações e Dados ---

# Chave de API do PVWatts (a chave original será substituída por uma variável de ambiente ou um placeholder seguro)
PVWATTS_API_KEY = os.environ.get('PVWATTS_API_KEY', 'hhMhbCRnyXob8EjS7U2AtoTVOjghmjbwDMfDa0Pu')
PVWATTS_URL = 'https://developer.nrel.gov/api/pvwatts/v8.json'
FILE_PATH_EQUIPAMENTOS = 'BDFotovoltaica.xlsx'

# --- Funções de Geocodificação ---

def geocode_location(location_name):
    """
    Converte um nome de local em coordenadas (latitude, longitude) usando Nominatim.
    """
    geolocator = Nominatim(user_agent="pv_dimensioning_app")
    try:
        location = geolocator.geocode(location_name, timeout=10)
        if location:
            return location.latitude, location.longitude
        return None, None
    except GeocoderTimedOut:
        print("Erro de Geocodificação: Tempo esgotado.")
        return None, None
    except GeocoderServiceError as e:
        print(f"Erro de Geocodificação: Serviço indisponível ou erro de requisição: {e}")
        return None, None
    except Exception as e:
        print(f"Erro inesperado na Geocodificação: {e}")
        return None, None

# --- Funções de Geração (PVWatts) ---

def fazer_requisicao_pvwatts(params):
    """Faz a requisição à API do PVWatts."""
    params['api_key'] = PVWATTS_API_KEY
    try:
        response = requests.get(PVWATTS_URL, params=params)
        response.raise_for_status()  # Levanta um erro para códigos de status HTTP ruins
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f'Erro na solicitação PVWatts: {e}')
        return None

def calcular_potencia_pico_necessaria(consumo_medio_mensal, latitude, longitude, azimuth, tilt, module_type=0, losses=14, array_type=1):
    """
    Calcula a potência de pico necessária (em kW) para atender ao consumo médio mensal.
    Retorna a potência de pico (kW) e os dados completos da simulação de 1kW.
    """
    consumo_anual = consumo_medio_mensal * 12
    
    # Parâmetros para simulação de 1 kW
    params = {
        'system_capacity': 1,  # 1 kW para cálculo de produção específica
        'module_type': module_type,
        'losses': losses,
        'array_type': array_type,
        'tilt': tilt,
        'azimuth': azimuth,
        'lat': latitude,
        'lon': longitude
    }
    
    data = fazer_requisicao_pvwatts(params)
    if data and 'outputs' in data and 'ac_annual' in data['outputs']:
        annual_energy_production_per_kw = data['outputs']['ac_annual']
        if annual_energy_production_per_kw > 0:
            potencia_pico_necessaria = consumo_anual / annual_energy_production_per_kw
            return potencia_pico_necessaria, data
    return None, None

def calcular_energia_gerada(potencia_pico, latitude, longitude, azimuth, tilt, module_type=0, losses=14, array_type=1):
    """
    Calcula a energia anual gerada (em kWh) para uma dada potência de pico.
    """
    params = {
        'system_capacity': potencia_pico,  # Potência de pico em kW
        'module_type': module_type,
        'losses': losses,
        'array_type': array_type,
        'tilt': tilt,
        'azimuth': azimuth,
        'lat': latitude,
        'lon': longitude,
        'dataset': 'intl'
    }
    
    data = fazer_requisicao_pvwatts(params)
    if data and 'outputs' in data and 'ac_annual' in data['outputs']:
        energy_production = data['outputs']['ac_annual']
        return energy_production
    return None

# --- Funções de Dimensionamento (Inversor e Arranjo) e Dados ---

def carregar_dados_equipamentos(file_path):
    """Carrega os dados de painéis e inversores do arquivo Excel."""
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
    """Salva os DataFrames de painéis e inversores no arquivo Excel."""
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df_paineis.to_excel(writer, sheet_name='paineis_solares', index=False)
            df_inversores.to_excel(writer, sheet_name='Inversores', index=False)
        return True
    except Exception as e:
        print(f"Erro ao salvar dados no Excel: {e}")
        return False

def salvar_novo_painel(novo_painel_data):
    """Adiciona um novo painel ao DataFrame e salva o arquivo Excel."""
    df_paineis, df_inversores = carregar_dados_equipamentos(FILE_PATH_EQUIPAMENTOS)
    
    # Criar um DataFrame com o novo painel
    novo_painel_df = pd.DataFrame([novo_painel_data])
    
    # Garantir que as colunas do novo painel correspondam às colunas existentes
    colunas_paineis = df_paineis.columns.tolist()
    novo_painel_df = novo_painel_df.reindex(columns=colunas_paineis, fill_value=None)
    
    # Concatenar e salvar
    df_paineis_atualizado = pd.concat([df_paineis, novo_painel_df], ignore_index=True)
    return salvar_dados_equipamentos(df_paineis_atualizado, df_inversores, FILE_PATH_EQUIPAMENTOS)

def salvar_novo_inversor(novo_inversor_data):
    """Adiciona um novo inversor ao DataFrame e salva o arquivo Excel."""
    df_paineis, df_inversores = carregar_dados_equipamentos(FILE_PATH_EQUIPAMENTOS)
    
    # Criar um DataFrame com o novo inversor
    novo_inversor_df = pd.DataFrame([novo_inversor_data])
    
    # Garantir que as colunas do novo inversor correspondam às colunas existentes
    colunas_inversores = df_inversores.columns.tolist()
    novo_inversor_df = novo_inversor_df.reindex(columns=colunas_inversores, fill_value=None)
    
    # Concatenar e salvar
    df_inversores_atualizado = pd.concat([df_inversores, novo_inversor_df], ignore_index=True)
    return salvar_dados_equipamentos(df_paineis, df_inversores_atualizado, FILE_PATH_EQUIPAMENTOS)

def selecionar_inversores(potencia_pico_w, df_inversores):
    """
    Seleciona combinações de inversores que atendam à potência de pico do sistema.
    A potência de pico (potencia_pico_w) deve estar em Watts.
    """
    potencia_inversor_necessaria = potencia_pico_w / 1.2
    
    combinacoes_inversores = []
    
    for _, inversor in df_inversores.iterrows():
        potencia_nominal_ca = inversor['maxima_potencia_nominal_ca']
        
        if pd.isna(potencia_nominal_ca) or potencia_nominal_ca <= 0:
            continue

        num_inversores_min = int(potencia_inversor_necessaria / potencia_nominal_ca)
        if potencia_inversor_necessaria % potencia_nominal_ca > 0:
             num_inversores_min += 1
        
        for num_inversores in range(num_inversores_min, num_inversores_min + 3):
            potencia_combinacao = num_inversores * potencia_nominal_ca
            
            potencia_pico_sistema = potencia_pico_w
            potencia_maxima_cc_suportada = potencia_combinacao * 1.2
            
            if potencia_maxima_cc_suportada >= potencia_pico_sistema * 0.95 and potencia_maxima_cc_suportada <= potencia_pico_sistema * 1.5:
                combinacoes_inversores.append({
                    'modelo': inversor['modelo'],
                    'fabricante': inversor['fabricante'],
                    'maxima_potencia_nominal_ca': potencia_nominal_ca,
                    'tensao_maxima_cc': inversor['tensao_maxima_cc'],
                    'tensao_start': inversor['tensao_start'],
                    'corrente_maxima_entrada_por_mpp_tracker': inversor['corrente_maxima_entrada_por_mpp_tracker'],
                    'numero_mpp_trackers': inversor['numero_mpp_trackers'],
                    'num_inversores': num_inversores,
                    'potencia_combinacao_ca': potencia_combinacao,
                    'potencia_maxima_cc_suportada': potencia_maxima_cc_suportada,
                    'diferenca_potencia_pico': abs(potencia_maxima_cc_suportada - potencia_pico_sistema)
                })
                
    df_combinacoes = pd.DataFrame(combinacoes_inversores)
    
    if df_combinacoes.empty:
        return df_combinacoes
        
    idx = df_combinacoes.groupby('modelo')['diferenca_potencia_pico'].idxmin()
    melhores_combinacoes = df_combinacoes.loc[idx].sort_values(by='diferenca_potencia_pico')
    
    return melhores_combinacoes

def calcular_arranjos_possiveis(df_paineis, tensao_maxima_cc, tensao_start):
    """
    Calcula todas as quantidades possíveis de módulos em série que atendem aos critérios de tensão do inversor.
    """
    arranjos = []
    
    for _, modulo in df_paineis.iterrows():
        potencia_modulo = modulo['potencia_maxima_nominal_pmax']
        tensao_voc = modulo['tensao_circuito_aberto_voc']
        corrente_isc = modulo['corrente_curto_circuito_isc']
        
        if pd.isna(tensao_voc) or tensao_voc <= 0:
            continue

        max_modulos_serie = int(tensao_maxima_cc / tensao_voc)
        
        for num_modulos_serie in range(1, max_modulos_serie + 1):
            tensao_total = num_modulos_serie * tensao_voc
            
            if tensao_start <= tensao_total <= tensao_maxima_cc:
                arranjos.append({
                    'modelo_painel': modulo['modelo'],
                    'fabricante_painel': modulo['fabricante'],
                    'potencia_modulo': potencia_modulo,
                    'num_modulos_serie': num_modulos_serie,
                    'tensao_total_voc': tensao_total,
                    'corrente_isc': corrente_isc
                })
                
    return pd.DataFrame(arranjos)

def selecionar_arranjo_paineis(potencia_pico_inversor_w, df_paineis, tensao_maxima_cc, tensao_start, corrente_maxima_mppt, num_mppt):
    """
    Seleciona o arranjo ideal de painéis para um único MPPT de um inversor.
    A potência de pico do inversor (potencia_pico_inversor_w) deve estar em Watts.
    """
    
    potencia_necessaria_mppt = potencia_pico_inversor_w / num_mppt
    
    potencia_minima = potencia_necessaria_mppt * 0.9
    potencia_maxima = potencia_necessaria_mppt * 1.1
    
    arranjos_possiveis = calcular_arranjos_possiveis(df_paineis, tensao_maxima_cc, tensao_start)
    
    if arranjos_possiveis.empty:
        return None
        
    arranjos_validos = []
    
    for _, arranjo in arranjos_possiveis.iterrows():
        potencia_total_serie = arranjo['num_modulos_serie'] * arranjo['potencia_modulo']
        
        if potencia_total_serie <= 0:
            continue

        max_conjuntos_paralelo_corrente = int(corrente_maxima_mppt / arranjo['corrente_isc'])
        max_conjuntos_paralelo_potencia = int(potencia_maxima / potencia_total_serie)
        
        max_conjuntos_paralelo = min(max_conjuntos_paralelo_corrente, max_conjuntos_paralelo_potencia)
        
        for num_conjuntos_paralelo in range(1, max_conjuntos_paralelo + 1):
            potencia_total_arranjo = potencia_total_serie * num_conjuntos_paralelo
            corrente_total = num_conjuntos_paralelo * arranjo['corrente_isc']
            
            if potencia_minima <= potencia_total_arranjo <= potencia_maxima and corrente_total <= corrente_maxima_mppt:
                arranjos_validos.append({
                    'modelo_painel': arranjo['modelo_painel'],
                    'fabricante_painel': arranjo['fabricante_painel'],
                    'potencia_modulo': arranjo['potencia_modulo'],
                    'num_modulos_serie': arranjo['num_modulos_serie'],
                    'num_conjuntos_paralelo': num_conjuntos_paralelo,
                    'potencia_total_arranjo_mppt': potencia_total_arranjo,
                    'corrente_total_isc': corrente_total,
                    'diferenca_potencia_mppt': abs(potencia_total_arranjo - potencia_necessaria_mppt)
                })
                
    if not arranjos_validos:
        return None
        
    df_arranjos_validos = pd.DataFrame(arranjos_validos)
    arranjo_ideal = df_arranjos_validos.loc[df_arranjos_validos['diferenca_potencia_mppt'].idxmin()]
    
    return arranjo_ideal.to_dict()

def dimensionar_sistema(potencia_pico_w, df_paineis, df_inversores):
    """
    Realiza o dimensionamento completo do sistema (inversor e arranjo) para a potência de pico.
    """
    melhores_combinacoes = selecionar_inversores(potencia_pico_w, df_inversores)
    
    if melhores_combinacoes.empty:
        return None, "Nenhuma combinação de inversores válida encontrada."
        
    resultados_dimensionamento = []
    
    for _, combinacao in melhores_combinacoes.iterrows():
        
        potencia_pico_por_inversor_w = potencia_pico_w / combinacao['num_inversores']
        
        tensao_maxima_cc = combinacao['tensao_maxima_cc']
        tensao_start = combinacao['tensao_start']
        corrente_maxima_mppt = combinacao['corrente_maxima_entrada_por_mpp_tracker']
        num_mppt = combinacao['numero_mpp_trackers']
        
        arranjo_ideal = selecionar_arranjo_paineis(
            potencia_pico_por_inversor_w, 
            df_paineis, 
            tensao_maxima_cc, 
            tensao_start, 
            corrente_maxima_mppt, 
            num_mppt
        )
        
        if arranjo_ideal:
            potencia_total_arranjo_mppt = arranjo_ideal['potencia_total_arranjo_mppt']
            potencia_total_sistema_w = potencia_total_arranjo_mppt * num_mppt * combinacao['num_inversores']
            
            num_modulos_serie = arranjo_ideal['num_modulos_serie']
            num_conjuntos_paralelo = arranjo_ideal['num_conjuntos_paralelo']
            num_total_paineis = num_modulos_serie * num_conjuntos_paralelo * num_mppt * combinacao['num_inversores']
            
            resultados_dimensionamento.append({
                'inversor_modelo': combinacao['modelo'],
                'inversor_fabricante': combinacao['fabricante'],
                'inversor_potencia_ca_nominal': combinacao['maxima_potencia_nominal_ca'],
                'inversor_num_unidades': combinacao['num_inversores'],
                'inversor_num_mppt': combinacao['numero_mpp_trackers'],
                'painel_modelo': arranjo_ideal['modelo_painel'],
                'painel_fabricante': arranjo_ideal['fabricante_painel'],
                'painel_potencia': arranjo_ideal['potencia_modulo'],
                'arranjo_modulos_serie': num_modulos_serie,
                'arranjo_conjuntos_paralelo_por_mppt': num_conjuntos_paralelo,
                'arranjo_potencia_total_mppt_w': potencia_total_arranjo_mppt,
                'sistema_potencia_total_w': potencia_total_sistema_w,
                'sistema_num_total_paineis': num_total_paineis,
                'sistema_potencia_pico_alvo_w': potencia_pico_w
            })
            
    if not resultados_dimensionamento:
        return None, "Nenhum arranjo de painéis válido encontrado para as combinações de inversores."
        
    return pd.DataFrame(resultados_dimensionamento), None

# --- Função Principal de Integração ---

def realizar_dimensionamento_completo(consumo_medio_mensal, latitude, longitude, azimuth, tilt):
    """
    Integra o cálculo de geração e o dimensionamento do sistema.
    """
    # 1. Calcular a potência de pico necessária (em kW)
    potencia_pico_kw, dados_pvwatts = calcular_potencia_pico_necessaria(
        consumo_medio_mensal, latitude, longitude, azimuth, tilt
    )
    
    if not potencia_pico_kw:
        return None, "Falha ao calcular a potência de pico necessária via PVWatts. Verifique as coordenadas e a conexão com a API."
        
    potencia_pico_w = potencia_pico_kw * 1000 # Converter para Watts
    
    # 2. Carregar dados de equipamentos
    df_paineis, df_inversores = carregar_dados_equipamentos(FILE_PATH_EQUIPAMENTOS)
    
    if df_paineis.empty or df_inversores.empty:
        return None, "Falha ao carregar dados de painéis ou inversores. Verifique o arquivo de equipamentos."
        
    # 3. Dimensionar o sistema (inversor e arranjo)
    df_dimensionamento, erro_dimensionamento = dimensionar_sistema(
        potencia_pico_w, df_paineis, df_inversores
    )
    
    if erro_dimensionamento:
        return None, erro_dimensionamento
        
    # 4. Calcular a energia gerada pelo sistema dimensionado (em kWh)
    energia_gerada_anual = calcular_energia_gerada(
        potencia_pico_kw, latitude, longitude, azimuth, tilt
    )
    
    # Adicionar informações de geração aos resultados
    df_dimensionamento['consumo_anual_kwh'] = consumo_medio_mensal * 12
    df_dimensionamento['potencia_pico_necessaria_kw'] = potencia_pico_kw
    df_dimensionamento['energia_gerada_anual_kwh'] = energia_gerada_anual
    
    return df_dimensionamento, None

if __name__ == '__main__':
    # Exemplo de uso para teste
    consumo = 1000 # kWh/mês
    lat = -20.46
    lon = -54.62
    azimuth = 180
    tilt = abs(lat)
    
    # Teste de Geocodificação
    print("--- Teste de Geocodificação ---")
    lat_geo, lon_geo = geocode_location("São Paulo, Brasil")
    print(f"São Paulo: Lat={lat_geo}, Lon={lon_geo}")
    
    # Teste de Dimensionamento
    resultados, erro = realizar_dimensionamento_completo(consumo, lat, lon, azimuth, tilt)
    
    if erro:
        print(f"Erro: {erro}")
    else:
        print("\n--- Resultados do Dimensionamento ---")
        print(resultados.to_markdown(index=False))
        
        # Exemplo de cálculo de potência de pico
        potencia_pico_kw, _ = calcular_potencia_pico_necessaria(consumo, lat, lon, azimuth, tilt)
        print(f"\nPotência de Pico Necessária: {potencia_pico_kw:.2f} kW")
        
        # Exemplo de cálculo de energia gerada
        energia_gerada = calcular_energia_gerada(potencia_pico_kw, lat, lon, azimuth, tilt)
        print(f"Energia Gerada (Alvo): {energia_gerada:.2f} kWh/ano")
