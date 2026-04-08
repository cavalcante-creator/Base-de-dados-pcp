import streamlit as st
import pandas as pd
import os
from datetime import datetime
import pytz

st.set_page_config(page_title="Base de Dados PCP", layout="wide")

# ==========================================================
# FUNÇÃO DATA/HORA
# ==========================================================

def agora():
    fuso = pytz.timezone("America/Sao_Paulo")
    data_atual = datetime.now(fuso)
    return data_atual

# ==========================================================
# MENU LATERAL
# ==========================================================

pagina = st.sidebar.radio(
    "Escolha uma opção:",
    [
        "Upload Ordens de Fabricação",
        "Upload Previsão Produção",
        "Upload Perfil e Saldos",
        "Upload Parâmetros",
        "Gerar Excel"
    ]
)

# ==========================================================
# UPLOAD ORDENS DE FABRICAÇÃO
# ==========================================================

if pagina == "Upload Ordens de Fabricação":

    st.title("📦 Upload Ordens de Fabricação")

    arquivo_of = st.file_uploader(
        "Selecione o arquivo de OF",
        type=["xlsx", "xls", "csv"]
    )

    if arquivo_of:
        st.success("Arquivo de OF carregado com sucesso!")

# ==========================================================
# UPLOAD PREVISÃO PRODUÇÃO
# ==========================================================

if pagina == "Upload Previsão Produção":

    st.title("📈 Upload Previsão Produção")

    arquivo_prev = st.file_uploader(
        "Selecione o arquivo de previsão",
        type=["xlsx", "xls", "csv"]
    )

    if arquivo_prev:
        st.success("Arquivo de previsão carregado com sucesso!")

# ==========================================================
# UPLOAD PERFIL E SALDOS
# ==========================================================

if pagina == "Upload Perfil e Saldos":

    st.title("📋 Upload Perfil de Itens / Saldos")

    arquivo_saldos = st.file_uploader(
        "Selecione o arquivo de Perfil de Itens",
        type=["xlsx", "xls", "csv"]
    )

    if arquivo_saldos:

        try:

            if arquivo_saldos.name.endswith(".csv"):
                df_saldos = pd.read_csv(arquivo_saldos)

            else:
                try:
                    arquivo_saldos.seek(0)
                    df_saldos = pd.read_excel(
                        arquivo_saldos,
                        engine="openpyxl"
                    )
                except:
                    arquivo_saldos.seek(0)
                    df_saldos = pd.read_excel(
                        arquivo_saldos,
                        engine="xlrd"
                    )

            df_saldos.columns = (
                df_saldos.columns
                .astype(str)
                .str.upper()
                .str.strip()
            )

            colunas_saldo = [
                "CODIGO",
                "DESCRICAO",
                "SALDO ALMOX 1",
                "SALDO ALMOX 2",
                "SALDO ALMOX 3",
                "SALDO ALMOX 4",
                "SALDO TOTAL"
            ]

            colunas_existentes = [
                c for c in colunas_saldo if c in df_saldos.columns
            ]

            if len(colunas_existentes) == 0:
                st.error("Nenhuma coluna de saldo encontrada.")

            else:

                df_saldos = df_saldos[colunas_existentes]

                df_saldos["Data Processamento"] = agora().strftime("%d/%m/%Y")
                df_saldos["Hora Processamento"] = agora().strftime("%H:%M:%S")

                arquivo_csv = "saldos.csv"

                if os.path.exists(arquivo_csv):
                    df_antigo = pd.read_csv(arquivo_csv)
                    df_saldos = pd.concat(
                        [df_antigo, df_saldos],
                        ignore_index=True
                    )

                df_saldos.to_csv(arquivo_csv, index=False)

                st.success("Saldos carregados com sucesso!")
                st.dataframe(df_saldos, use_container_width=True)

        except Exception as e:
            st.error(f"Erro ao processar arquivo: {e}")

# ==========================================================
# UPLOAD PARÂMETROS
# ==========================================================

if pagina == "Upload Parâmetros":

    st.title("⚙️ Parâmetros Produção")

    file = st.file_uploader(
        "Excel Parâmetros",
        type=["xls", "xlsx"],
        accept_multiple_files=False
    )

    if file:

        try:
            file.seek(0)
            df_raw = pd.read_excel(
                file,
                header=None,
                engine="openpyxl"
            )
            engine_excel = "openpyxl"

        except:
            file.seek(0)
            df_raw = pd.read_excel(
                file,
                header=None,
                engine="xlrd"
            )
            engine_excel = "xlrd"

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

                st.success("Parâmetros carregados com sucesso!")
                st.dataframe(df, use_container_width=True)

# ==========================================================
# GERAR EXCEL
# ==========================================================

if pagina == "Gerar Excel":

    st.title("📥 Gerar Excel")

    if os.path.exists("parametros.csv"):
        df_parametros = pd.read_csv("parametros.csv")

        st.subheader("Parâmetros")
        st.dataframe(df_parametros, use_container_width=True)

        st.download_button(
            label="Baixar Parâmetros CSV",
            data=df_parametros.to_csv(index=False).encode("utf-8"),
            file_name="parametros_exportados.csv",
            mime="text/csv"
        )

    if os.path.exists("saldos.csv"):
        df_saldos = pd.read_csv("saldos.csv")

        st.subheader("Saldos")
        st.dataframe(df_saldos, use_container_width=True)

        st.download_button(
            label="Baixar Saldos CSV",
            data=df_saldos.to_csv(index=False).encode("utf-8"),
            file_name="saldos_exportados.csv",
            mime="text/csv"
        )

    if not os.path.exists("parametros.csv") and not os.path.exists("saldos.csv"):
        st.warning("Nenhum arquivo foi carregado ainda.")
