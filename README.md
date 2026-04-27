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
