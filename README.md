# Aurumcalc - Dimensionamento Fotovoltaico

## Visão geral
O app mantém o fluxo clássico de dimensionamento FV com PVWatts + seleção de inversores/módulos por datasheet (Excel), e adiciona a seção **Análise Horária e Tarifária** para avaliação em passo de 1h.

## Nova camada de perfis de carga
Estrutura em `data/load_profiles/`:
- `profiles_index.json`
- subpastas por setor (`escola/`, `supermercado/`, `frigorifico/`, `industria/`, `escritorio/`, `varejo/`, `restaurante/`, `hotel/`, `hospital/`, `custom/`)

Cada perfil JSON possui metadados e vetor horário (`hourly_vector_normalized` ou `hourly_vector_kw`).

> Perfis com `tipo_perfil = synthetic_for_testing` são **apenas para teste** e não devem ser usados para proposta real sem substituir por base validada/medida.

## Formato esperado do perfil (JSON)
Campos recomendados:
- `id`, `nome`, `setor`, `subtipo`
- `pais_origem`, `fonte`
- `resolucao_original`, `resolucao_final`
- `ano_climatico_ou_origem`, `unidade_original`
- `tipo_perfil` (`measured`, `simulated`, `synthetic_for_testing`)
- `observacoes`
- `hourly_vector_normalized` (24 ou 8760) **ou** `hourly_vector_kw` (24 ou 8760)

## Calibração por fator único
Módulo: `hourly_analysis.calibrate_load_profile`.
- Usa curva normalizada candidata.
- Expande para mês de análise.
- Aplica máscara de funcionamento (dias/horários).
- Separa ponta/fora ponta por calendário horário.
- Calcula **um único fator multiplicativo** (não separa ponta/fora ponta).
- Retorna curva calibrada, perfil escolhido, score, erros percentuais e alertas.

### Limitações da calibração
- Um único fator pode não representar bem clientes com múltiplos regimes operacionais.
- Erro alto indica necessidade de trocar perfil base ou usar curva medida.

## Grupo A x Grupo B
Modelos separados:
- `BillingGroupAInput`: energia/demanda em ponta e fora ponta, tarifas de energia e demanda.
- `BillingGroupBInput`: energia total e tarifas (convencional/branca).

## Interpretação da análise horária
A análise calcula por hora:
- carga
- geração FV
- autoconsumo
- importação da rede
- exportação

Depois separa por posto tarifário e estima economia:
1. autoconsumo instantâneo;
2. compensação de excedente (quando parametrizada);
3. redução de demanda;
4. load shifting;
5. peak shaving.

## Testes
Arquivo: `tests/test_hourly_analysis.py`.
Cobertura inicial:
- classificação ponta/fora ponta e fim de semana/feriado;
- calibração com fator único;
- balanço horário (autoconsumo/importação/exportação);
- separação por posto tarifário e demanda máxima;
- limite por área;
- load shifting simples;
- peak shaving simples.

## Novo modo: Dimensionamento com Curva de Demanda Ajustada
O app mantém o fluxo convencional **inalterado** e adiciona o fluxo avançado:

1. seleção de tipo de cliente (curva típica sintética de engenharia);
2. ajuste da curva horária pela conta (Grupo A ou Grupo B);
3. definição de área disponível por mapa (polígono);
4. aplicação de restrição de área no número máximo de módulos;
5. chamada do mesmo motor atual (PVWatts + seleção de módulos/inversores por datasheet);
6. análises complementares: autoconsumo, importação/exportação, demanda líquida, peak shaving e load shifting simplificado.

### Dados necessários
- **Grupo B**: `energia_total_kwh_mes`, `dias_funcionamento_mes`, `tipo_cliente` e (opcional) tarifa/demanda estimada.
- **Grupo A**: energia ponta/fora ponta, demanda ponta/fora ponta, dias de funcionamento, horário de ponta (default 18:00–21:00, editável) e tarifas opcionais.

### Ajuste da curva
- Função principal: `fit_load_profile_to_bill(...)`.
- Grupo B: escala a curva para fechar energia mensal.
- Grupo A: ajusta fatores separados de ponta e fora ponta e emite alertas quando a demanda estimada diverge da medida.

### Restrição por área de mapa
- Módulo: `src/map_area.py`.
- O usuário informa endereço/coordenadas, desenha polígono e obtém área em m².
- O limite de módulos usa:
  `max_modules_by_area = floor(available_area_m2 * packing_factor / module_area_m2)`.
- `packing_factor` editável entre 0.50 e 0.90 (default 0.70).

### Limitações atuais
- Curvas típicas em `data/load_profiles/typical_load_profiles.json` são sintéticas e destinadas a estudos preliminares.
- O load shifting desta versão é simplificado (percentual deslocável na ponta), sem otimização multiobjetivo ou modelagem de bateria.
- A área em mapa depende da precisão do desenho do usuário e da geocodificação.

### Próximos passos recomendados
- Calibrar perfis com memória de massa / medição real.
- Incorporar sazonalidade mensal e múltiplos regimes operacionais.
- Expandir modelo econômico com impostos, compensação e demanda contratada detalhada.
- Incluir otimização de load shifting com restrições operacionais reais.
