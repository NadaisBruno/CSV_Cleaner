import streamlit as st  # User Interface do Streamlit(substitui o tkinter com um layout bem mais moderno)
import pandas as pd
from csv_cleaner import ler_csv, limpar_dados, gerar_relatorio, exportar_excel

st.title("CSV Cleaner - Versao Streamlit")  # titulo da pagina
st.write("""
### Como usar:
1️⃣ Carregue um ficheiro CSV
2️⃣ Clique em **Limpar dados**
3️⃣ Clique em Gerar relatório
4️⃣ Faça o download do Excel final
""")  # apresenta argumentos na pagina

# ------------------Upload Do Ficheiro---------------------------- #

uploaded_file = st.file_uploader("Escolher o ficheiro CSV", type=["csv"])  # o utilizador escolhe o ficheiro // 'type' é a extensao permitida do ficheiro escolhido
if uploaded_file is not None:
    # st.write("DEBUG — ficheiro recebido:", uploaded_file.name)
    st.success("Ficheiro carregado com sucesso!")

    # ler o CSV
    df = ler_csv(uploaded_file)

    # st.write("DEBUG — df:", df)

    # Verificacao de seguranca: Se o utilizador ainda não carregou nenhum CSV, a variável df estará igual a None.
    # Se escrevermos st.dataframe(df) sem este "if", davas um erro no Streamlit.
    if df is not None:  # só executa o que vem a seguir se o CSV foi lido com sucesso
        st.write("## Dados brutos")  # mostra o titulo
        st.dataframe(df)  # mostramos o dataframe pandas no Streamlit(serve para o utilizador ver csv original)
        # criamos o botao para limpar dados
        if st.button("Limpar_dados"):
            # st.write("DEBUG — botão clicado!") # so para teste - sera removido depois
            st.session_state["df_limpo"] = limpar_dados(df)  # chama a def limpar_dados e devolve o novo dataframe limpo, ou seja, guarda o dataframe limpo na memoria do streamlit para ser usado depois

        if "df_limpo" in st.session_state:  # verifica se o utilizador clicou no botao
            st.subheader("Dados Limpos")  # mostra o subtitulo
            st.dataframe(st.session_state["df_limpo"])  # vai buscar os dados limpos a memória e mostra-os no navegador

            # criar botao gerar relatorio.Este botao so aparece porque o df_limpo ja existe
            if st.button("Gerar Relatório"):
                # chamamos a def que cria o relatorio
                relatorio = gerar_relatorio(st.session_state["df_limpo"])
                # guardamos o relatorio na memória
                st.session_state["relatorio"] = relatorio

                if "relatorio" in st.session_state:  # se o relatorio foi guardado na memória
                    st.success("Relatorio criado com sucesso")
                    st.info("Pode agora exportar o ficheiro Excel")
                    if st.button("Exportar Excel"):  # cria o botao e se o utilizador clicar no botão Exportar Excel,
                        # então pega nos dados limpos e no relatório, e cria o ficheiro Excel
                        exportar_excel(
                            st.session_state["df_limpo"],
                            st.session_state["relatorio"]
                        )
                    st.success("Excel gerado.Agora posso fazer o download.")

                    # Abrir o ficheiro e disponibilizar o ‘download’ no navegador
                    # rb = read binary porque o excel nao é texto, é um ficheiro binario
                    with open("relatorio.xlsx", "rb") as f:
                        # criamos o botao para fazer o download do ficheiro
                        st.download_button(
                            label="📥Download do Excel",
                            data=f,
                            file_name="relatorio.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )