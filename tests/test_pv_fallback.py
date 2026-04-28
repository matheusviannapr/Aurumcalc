import pandas as pd

import pv_calculator as pvc


def test_realizar_dimensionamento_usa_fallback_quando_pvwatts_falha(monkeypatch):
    monkeypatch.setattr(pvc, "calcular_potencia_pico_necessaria", lambda *args, **kwargs: (None, None, None))

    df_p = pd.DataFrame(
        [
            {
                "modelo": "painel_teste",
                "fabricante": "fab",
                "potencia_maxima_nominal_pmax": 550,
                "tensao_operacao_otima_vmp": 41,
                "corrente_operacao_otima_imp": 13,
                "tensao_circuito_aberto_voc": 50,
                "corrente_curto_circuito_isc": 14,
                "eficiencia_modulo": 21,
            }
        ]
    )
    df_i = pd.DataFrame(
        [
            {
                "modelo": "inv_teste",
                "fabricante": "fab",
                "potencia_maxima_fv_maxima": 20000,
                "tensao_maxima_cc": 1100,
                "tensao_start": 200,
                "tensao_nominal": 600,
                "faixa_tensao_mpp": "200-850",
                "numero_mpp_trackers": 2,
                "corrente_maxima_entrada_por_mppt_tracker": 30,
                "corrente_maxima_curto_circuito_por_mppt_tracker": 40,
                "maxima_potencia_nominal_ca": 15000,
                "potencia_maxima_aparente_ca": 16000,
                "tensao_nominal_ca": 380,
                "frequencia_rede_ca": "60Hz",
                "corrente_saida_maxima": 30,
                "fator_potencia_ajustavel": "0.8i-0.8c",
                "quantidade_fases_ca": 3,
            }
        ]
    )

    monkeypatch.setattr(pvc, "carregar_dados_equipamentos", lambda *args, **kwargs: (df_p, df_i))
    monkeypatch.setattr(pvc, "calcular_energia_gerada", lambda *args, **kwargs: (None, None))
    monkeypatch.setattr(
        pvc,
        "dimensionar_sistema",
        lambda *args, **kwargs: (
            pd.DataFrame(
                [
                    {
                        "inversor_modelo": "inv_teste",
                        "inversor_fabricante": "fab",
                        "inversor_num_unidades": 1,
                        "painel_modelo": "painel_teste",
                        "painel_fabricante": "fab",
                        "painel_potencia": 550,
                        "arranjo_modulos_serie": 10,
                        "arranjo_conjuntos_paralelo_por_mppt": 2,
                        "inversor_num_mppt": 2,
                        "sistema_num_total_paineis": 40,
                        "sistema_potencia_total_w": 22000,
                    }
                ]
            ),
            None,
        ),
    )

    out, erro = pvc.realizar_dimensionamento_completo(
        consumo_medio_mensal=5000,
        latitude=-23.55,
        longitude=-46.63,
        azimuth=180,
        tilt=23,
    )

    assert erro is None
    assert out is not None and not out.empty
    assert out.iloc[0]["fonte_geracao"] == "estimativa_fallback"
    assert float(out.iloc[0]["potencia_pico_necessaria_kw"]) > 0
