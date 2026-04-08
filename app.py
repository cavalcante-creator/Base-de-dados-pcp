# ==========================================================
# PARÂMETROS
# ==========================================================

if pagina == "Upload Parâmetros":

    st.title("⚙️ Parâmetros Produção")

    file = st.file_uploader(
        "Excel Parâmetros",
        type=["xls", "xlsx"],
        accept_multiple_files=False
    )

    if file:

        if file.name.endswith(".xls"):
            engine_excel = "xlrd"
        else:
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

            colunas_encontradas = [
                c for c in colunas_necessarias if c in df.columns
            ]

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
