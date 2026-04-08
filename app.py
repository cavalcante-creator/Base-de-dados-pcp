# ==========================================================
# PREVISÃO
# ==========================================================

if pagina == "Upload Previsão Produção":

    st.title("📅 Previsão Produção")

    file = st.file_uploader("Excel Previsão", type=["xlsx"])

    if file:

        df_raw = pd.read_excel(file, header=None, engine="openpyxl")

        linha_header = None
        for i in range(len(df_raw)):
            if "COD" in df_raw.iloc[i].astype(str).str.upper().values:
                linha_header = i
                break

        file.seek(0)

        df = pd.read_excel(file, header=linha_header, engine="openpyxl")
        df.columns = df.columns.astype(str).str.upper()

        col_cod = [c for c in df.columns if "COD" in c][0]
        col_prod = [c for c in df.columns if "PROD" in c][0]
        col_prev = [c for c in df.columns if "PREVIS" in c][0]

        df = df[[col_cod, col_prod, col_prev]]
        df.columns = ["COD", "PRODUTO", "PREVISAO"]

        df["PREVISAO"] = pd.to_numeric(df["PREVISAO"], errors="coerce").fillna(0)

        df["Data Processamento"] = agora().strftime("%d/%m/%Y")
        df["Hora Processamento"] = agora().strftime("%H:%M:%S")

        arquivo = "previsao.csv"

        if os.path.exists(arquivo):
            df_antigo = pd.read_csv(arquivo)
            df = pd.concat([df_antigo, df], ignore_index=True)

        df.to_csv(arquivo, index=False)

        st.success("Previsão carregada!")
        st.dataframe(df, use_container_width=True)

# ==========================================================
# PARÂMETROS
# ==========================================================

if pagina == "Upload Parâmetros":

    st.title("⚙️ Parâmetros Produção")

    file = st.file_uploader("Excel Parâmetros", type=["xlsx"])

    if file:

        engine_excel = "openpyxl"

        df_raw = pd.read_excel(
            file,
            header=None,
            engine=engine_excel
        )

        linha_header = None
        for i in range(len(df_raw)):
            linha = df_raw.iloc[i].astype(str).str.upper().str.strip().values

            if "COD ITEM" in linha:
                linha_header = i
                break

        if linha_header is None:
            st.error("Não foi possível localizar a linha de cabeçalho.")
        else:

            file.seek(0)

            df = pd.read_excel(
                file,
                header=linha_header,
                engine=engine_excel
            )

            df.columns = df.columns.astype(str).str.upper().str.strip()

            colunas_necessarias = [
                "COD ITEM",
                "DESC TECNICA",
                "UM",
                "LOTE MIN",
                "LOTE MAX",
                "LOTE MULT",
                "ESTQ SEG",
                "TEMP REP",
                "TEMP SEG",
                "AGRUP",
                "PLANEJADOR",
                "CONS MEDIO"
            ]

            colunas_encontradas = [c for c in colunas_necessarias if c in df.columns]

            if len(colunas_encontradas) == 0:
                st.error("Nenhuma das colunas esperadas foi encontrada no arquivo.")
            else:

                df = df[colunas_encontradas]

                renomear = {
                    "COD ITEM": "COD ITEM",
                    "DESC TECNICA": "DESCRICAO",
                    "UM": "UM",
                    "LOTE MIN": "LOTE MIN",
                    "LOTE MAX": "LOTE MAX",
                    "LOTE MULT": "LOTE MULT",
                    "ESTQ SEG": "ESTOQUE SEGURANCA",
                    "TEMP REP": "TEMPO REPOSICAO",
                    "TEMP SEG": "TEMPO SEGURANCA",
                    "AGRUP": "AGRUPAMENTO",
                    "PLANEJADOR": "PLANEJADOR",
                    "CONS MEDIO": "CONSUMO MEDIO"
                }

                df = df.rename(columns=renomear)

                df["Data Processamento"] = agora().strftime("%d/%m/%Y")
                df["Hora Processamento"] = agora().strftime("%H:%M:%S")

                arquivo = "parametros.csv"

                if os.path.exists(arquivo):
                    df_antigo = pd.read_csv(arquivo)
                    df = pd.concat([df_antigo, df], ignore_index=True)

                df.to_csv(arquivo, index=False)

                st.success("Parâmetros carregados!")
                st.dataframe(df, use_container_width=True)

# ==========================================================
# GERAR EXCEL
# ==========================================================

if pagina == "Gerar Excel":

    st.title("📥 Gerar Excel")

    arquivos = {
        "Saldo": "saldo.csv",
        "Perfil": "perfil.csv",
        "Ordens": "ordens.csv",
        "Previsão": "previsao.csv",
        "Parâmetros": "parametros.csv"
    }

    for nome, arquivo in arquivos.items():

        if os.path.exists(arquivo):

            df = pd.read_csv(arquivo)

            if "Data Processamento" in df.columns:
                df = df.sort_values(
                    by=["Data Processamento", "Hora Processamento"],
                    ascending=False
                )

            st.subheader(nome)
            st.dataframe(df, use_container_width=True)

            if st.button(f"Gerar Excel {nome}"):

                nome_excel = f"{nome}_{agora().strftime('%H-%M-%S')}.xlsx"

                df.to_excel(nome_excel, index=False)

                with open(nome_excel, "rb") as f:
                    st.download_button(
                        f"📥 Baixar {nome}",
                        f,
                        file_name=nome_excel
                    )
