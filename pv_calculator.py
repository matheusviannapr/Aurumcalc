import pandas as pd
import requests
import os

# --- Configurações e Dados ---

# Chave de API do PVWatts (a chave original será substituída por uma variável de ambiente ou um placeholder seguro)
# A chave fornecida no código original (hhMhbCRnyXob8EjS7U2AtoTVOjghmjbwDMfDa0Pu) será usada como placeholder
# Em um ambiente de produção, esta chave deve ser gerenciada de forma segura.
PVWATTS_API_KEY = os.environ.get('PVWATTS_API_KEY', 'hhMhbCRnyXob8EjS7U2AtoTVOjghmjbwDMfDa0Pu')
PVWATTS_URL = 'https://developer.nrel.gov/api/pvwatts/v8.json'

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
        'lon': longitude
    }
    
    data = fazer_requisicao_pvwatts(params)
    if data and 'outputs' in data and 'ac_annual' in data['outputs']:
        energy_production = data['outputs']['ac_annual']
        return energy_production
    return None

# --- Funções de Dimensionamento (Inversor e Arranjo) ---

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

def selecionar_inversores(potencia_pico_w, df_inversores):
    """
    Seleciona combinações de inversores que atendam à potência de pico do sistema.
    A potência de pico (potencia_pico_w) deve estar em Watts.
    """
    # Converter potência de pico de kW para W (se necessário, mas o código original usa W)
    # O código original usa 180000 W como exemplo, então vamos manter a entrada em W.
    
    # Ajuste para 20% de overload (1.2)
    potencia_inversor_necessaria = potencia_pico_w / 1.2
    
    combinacoes_inversores = []
    
    # Iterar sobre cada modelo de inversor
    for _, inversor in df_inversores.iterrows():
        # A potência nominal CA máxima do inversor
        potencia_nominal_ca = inversor['maxima_potencia_nominal_ca']
        
        if pd.isna(potencia_nominal_ca) or potencia_nominal_ca <= 0:
            continue # Ignorar inversores sem potência nominal válida

        # Calcular o número mínimo de inversores
        num_inversores_min = int(potencia_inversor_necessaria / potencia_nominal_ca)
        if potencia_inversor_necessaria % potencia_nominal_ca > 0:
             num_inversores_min += 1
        
        # Tentar combinações a partir do mínimo necessário até um limite razoável (ex: +2)
        for num_inversores in range(num_inversores_min, num_inversores_min + 3):
            potencia_combinacao = num_inversores * potencia_nominal_ca
            
            # Critério de seleção: a potência combinada deve estar entre 90% e 110% da potência necessária
            # O código original usa o critério de que a potência combinada deve ser >= 90% da potência necessária
            # e a potência de pico do sistema deve ser <= 120% da potência combinada.
            # Vamos simplificar para: a potência combinada deve ser suficiente para a potência de pico (com 20% de overload)
            
            # Critério do código original:
            # while num_inversores * inversor['maxima_potencia_nominal_ca'] <= potencia_inversor_necessaria * 1.1:
            #     ...
            #     if potencia_combinacao >= potencia_inversor_necessaria * 0.9:
            
            # Vamos usar um critério mais direto: a potência combinada deve ser maior ou igual à potência necessária
            # e a potência de pico do sistema deve ser <= 120% da potência combinada.
            
            # Potência de pico do sistema (W)
            potencia_pico_sistema = potencia_pico_w
            
            # Potência máxima de entrada CC que a combinação de inversores suporta (20% de overload)
            potencia_maxima_cc_suportada = potencia_combinacao * 1.2
            
            if potencia_maxima_cc_suportada >= potencia_pico_sistema * 0.95 and potencia_maxima_cc_suportada <= potencia_pico_sistema * 1.5: # Margem de 5% a 50%
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
        
    # Selecionar a melhor combinação para cada modelo (a que tem a menor diferença de potência de pico)
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

        # Encontrar a quantidade máxima de módulos em série
        max_modulos_serie = int(tensao_maxima_cc / tensao_voc)
        
        for num_modulos_serie in range(1, max_modulos_serie + 1):
            tensao_total = num_modulos_serie * tensao_voc
            
            # Critério de tensão: Tensão de start <= Tensão total <= Tensão máxima CC
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
    
    # Potência necessária por MPPT (W)
    potencia_necessaria_mppt = potencia_pico_inversor_w / num_mppt
    
    # Margem de potência (90% a 110% da potência necessária por MPPT)
    potencia_minima = potencia_necessaria_mppt * 0.9
    potencia_maxima = potencia_necessaria_mppt * 1.1
    
    # Calcular todos os arranjos possíveis que atendem aos critérios de tensão
    arranjos_possiveis = calcular_arranjos_possiveis(df_paineis, tensao_maxima_cc, tensao_start)
    
    if arranjos_possiveis.empty:
        return None
        
    arranjos_validos = []
    
    for _, arranjo in arranjos_possiveis.iterrows():
        potencia_total_serie = arranjo['num_modulos_serie'] * arranjo['potencia_modulo']
        
        if potencia_total_serie <= 0:
            continue

        # Calcular quantos conjuntos em paralelo são necessários para satisfazer a potência total por MPPT
        # O número máximo de conjuntos em paralelo é limitado pela corrente máxima do MPPT
        max_conjuntos_paralelo_corrente = int(corrente_maxima_mppt / arranjo['corrente_isc'])
        
        # O número máximo de conjuntos em paralelo é limitado pela potência máxima
        max_conjuntos_paralelo_potencia = int(potencia_maxima / potencia_total_serie)
        
        max_conjuntos_paralelo = min(max_conjuntos_paralelo_corrente, max_conjuntos_paralelo_potencia)
        
        for num_conjuntos_paralelo in range(1, max_conjuntos_paralelo + 1):
            potencia_total_arranjo = potencia_total_serie * num_conjuntos_paralelo
            corrente_total = num_conjuntos_paralelo * arranjo['corrente_isc']
            
            # Critério de potência: Potência mínima <= Potência total do arranjo <= Potência máxima
            # Critério de corrente: Corrente total <= Corrente máxima do MPPT
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
        
    # Selecionar o arranjo que está mais próximo da potência necessária por MPPT
    df_arranjos_validos = pd.DataFrame(arranjos_validos)
    arranjo_ideal = df_arranjos_validos.loc[df_arranjos_validos['diferenca_potencia_mppt'].idxmin()]
    
    return arranjo_ideal.to_dict()

def dimensionar_sistema(potencia_pico_w, df_paineis, df_inversores):
    """
    Realiza o dimensionamento completo do sistema (inversor e arranjo) para a potência de pico.
    """
    # 1. Selecionar as melhores combinações de inversores
    melhores_combinacoes = selecionar_inversores(potencia_pico_w, df_inversores)
    
    if melhores_combinacoes.empty:
        return None, "Nenhuma combinação de inversores válida encontrada."
        
    resultados_dimensionamento = []
    
    # 2. Para cada melhor combinação, calcular o arranjo ideal
    for _, combinacao in melhores_combinacoes.iterrows():
        
        # Potência de pico do sistema dividida pelo número de inversores
        potencia_pico_por_inversor_w = potencia_pico_w / combinacao['num_inversores']
        
        # Parâmetros do inversor
        tensao_maxima_cc = combinacao['tensao_maxima_cc']
        tensao_start = combinacao['tensao_start']
        corrente_maxima_mppt = combinacao['corrente_maxima_entrada_por_mpp_tracker']
        num_mppt = combinacao['numero_mpp_trackers']
        
        # Selecionar arranjo de painéis para um MPPT
        arranjo_ideal = selecionar_arranjo_paineis(
            potencia_pico_por_inversor_w, 
            df_paineis, 
            tensao_maxima_cc, 
            tensao_start, 
            corrente_maxima_mppt, 
            num_mppt
        )
        
        if arranjo_ideal:
            # Calcular a potência total do sistema com o arranjo selecionado
            potencia_total_arranjo_mppt = arranjo_ideal['potencia_total_arranjo_mppt']
            potencia_total_sistema_w = potencia_total_arranjo_mppt * num_mppt * combinacao['num_inversores']
            
            # Calcular o número total de painéis
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

def realizar_dimensionamento_completo(consumo_medio_mensal, latitude, longitude, azimuth, tilt, file_path_equipamentos):
    """
    Integra o cálculo de geração e o dimensionamento do sistema.
    """
    # 1. Calcular a potência de pico necessária (em kW)
    potencia_pico_kw, dados_pvwatts = calcular_potencia_pico_necessaria(
        consumo_medio_mensal, latitude, longitude, azimuth, tilt
    )
    
    if not potencia_pico_kw:
        return None, "Falha ao calcular a potência de pico necessária via PVWatts."
        
    potencia_pico_w = potencia_pico_kw * 1000 # Converter para Watts
    
    # 2. Carregar dados de equipamentos
    df_paineis, df_inversores = carregar_dados_equipamentos(file_path_equipamentos)
    
    if df_paineis.empty or df_inversores.empty:
        return None, "Falha ao carregar dados de painéis ou inversores."
        
    # 3. Dimensionar o sistema (inversor e arranjo)
    df_dimensionamento, erro_dimensionamento = dimensionar_sistema(
        potencia_pico_w, df_paineis, df_inversores
    )
    
    if erro_dimensionamento:
        return None, erro_dimensionamento
        
    # 4. Calcular a energia gerada pelo sistema dimensionado (em kWh)
    # Usamos a potência total do sistema dimensionado (em W) e convertemos para kW
    df_dimensionamento['sistema_potencia_total_kw'] = df_dimensionamento['sistema_potencia_total_w'] / 1000
    
    # Vamos usar a potência de pico alvo para o cálculo final de geração, pois é o que o usuário deseja atender
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
    file_path = 'BDFotovoltaica.xlsx'
    consumo = 1000 # kWh/mês
    lat = -20.46
    lon = -54.62
    azimuth = 180
    tilt = abs(lat)
    
    resultados, erro = realizar_dimensionamento_completo(consumo, lat, lon, azimuth, tilt, file_path)
    
    if erro:
        print(f"Erro: {erro}")
    else:
        print("--- Resultados do Dimensionamento ---")
        print(resultados.to_markdown(index=False))
        
        # Exemplo de cálculo de potência de pico
        potencia_pico_kw, _ = calcular_potencia_pico_necessaria(consumo, lat, lon, azimuth, tilt)
        print(f"\nPotência de Pico Necessária: {potencia_pico_kw:.2f} kW")
        
        # Exemplo de cálculo de energia gerada
        energia_gerada = calcular_energia_gerada(potencia_pico_kw, lat, lon, azimuth, tilt)
        print(f"Energia Gerada (Alvo): {energia_gerada:.2f} kWh/ano")
