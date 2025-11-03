from memoria_calculo import gerar_memoria_calculo_latex

# ...
if st.button("Gerar Memória de Cálculo (LaTeX)"):
    try:
        caminho = gerar_memoria_calculo_latex(
            resultados_df=resultados_df,
            caminho_tex="memoria_calculo.tex",
            projeto="Dimensionamento FV — Cliente X",
            cliente="Cliente X",
            local="Cidade/UF",
            latitude=latitude,
            longitude=longitude,
            azimuth=azimuth,
            tilt=tilt,
            observacoes="Memória de cálculo gerada automaticamente a partir do dimensionamento."
        )
        st.success("Memória de cálculo gerada com sucesso!")
        st.download_button("Baixar .tex", data=open(caminho, "rb").read(), file_name="memoria_calculo.tex")
    except Exception as e:
        st.error(f"Não foi possível gerar a memória de cálculo LaTeX: {e}")
