import streamlit as st
import pandas as pd
import os
import io
import re
from datetime import datetime
import pytz

st.set_page_config(page_title="Base de Dados PCP", layout="wide")

# =========================
# FUNÇÕES
# =========================

def data_hora_brasil():
    fuso = pytz.timezone("America/Sao_Paulo")
    return datetime.now(fuso).strftime("%d/%m/%Y %H:%M:%S")

def identificar_linha_header(df_raw):
    for i in range(min(15, len(df_raw))):
        linha = df_raw.iloc[i].astype(str).str.upper()

        if any("COD" in str(v) for v in linha) and any("DESC" in str(v) for v in linha):
            return i

        if any("CÓDIGO" in str(v) for v in linha):
            return i

        if any("DESCRI" in str(v) for v in linha):
            return i

    return 0

def abrir_arquivo_generico(file):
    nome_arquivo = file.name.lower()

    try:
        if nome_arquivo.endswith(".csv"):
            file.seek(0)
            try:
                return pd.read_csv(file, sep=";", encoding="latin1", header=None)
            except:
                file.seek(0)
                return pd.read_csv(file, sep=",", encoding="latin1", header=None)

        elif nome_arquivo.endswith(".xlsx"):
            file.seek(0)
            return pd.read_excel(file, header=None, engine="openpyxl")

        elif nome_arquivo.endswith(".xls"):
            file.seek(0)

            try:
                return pd.read_excel(file, header=None, engine="xlrd")
            except:
                file.seek(0)

                try:
                    conteudo = file.read().decode("latin1", errors="ignore")

                    tabelas = pd.read_html(io.StringIO(conteudo))

                    if len(tabelas) > 0:
                        return tabelas[0]
                    else:
                        st.error("Não foi encontrada nenhuma tabela no arquivo.")
                        st.stop()

                except Exception as e:
                    st.error(f"Erro ao abrir arquivo .xls: {e}")
                    st.stop()

        else:
            st.error("Formato de arquivo não suportado.")
            st.stop()

    except Exception as e:
        st.error(f"Erro ao abrir arquivo: {e}")
        st.stop()

def carregar_dataframe(file):
    nome_arquivo = file.name.lower()

    try:
        if nome_arquivo.endswith(".csv"):
            file.seek(0)

            try:
                df_raw = pd.read_csv(file, sep=";", encoding="latin1", header=None)
            except:
                file.seek(0)
                df_raw = pd.read_csv(file, sep=",", encoding="latin1", header=None)

            linha_header = identificar_linha_header(df_raw)

            file.seek(0)

            try:
                df = pd.read_csv(file, sep=";", encoding="latin1", header=linha_header)
            except:
                file.seek(0)
                df = pd.read_csv(file, sep=",", encoding="latin1", header=linha_header)

            return df

        elif nome_arquivo.endswith(".xlsx"):
            file.seek(0)
            df_raw = pd.read_excel(file, header=None, engine="openpyxl")

            linha_header = identificar_linha_header(df_raw)

            file.seek(0)
            df = pd.read_excel(file, header=linha_header, engine="openpyxl")

            return df

        elif nome_arquivo.endswith(".xls"):
            file.seek(0)

            try:
                df_raw = pd.read_excel(file, header=None, engine="xlrd")
                linha_header = identificar_linha_header(df_raw)

                file.seek(0)
                df = pd.read_excel(file, header=linha_header, engine="xlrd")

                return df

            except:
                file.seek(0)

                try:
                    conteudo = file.read().decode("latin1", errors="ignore")
                    tabelas = pd.read_html(io.StringIO(conteudo))

                    if len(tabelas) == 0:
                        st.error("Nenhuma tabela encontrada no arquivo.")
                        st.stop()

                    df_raw = tabelas[0]
                    linha_header = identificar_linha_header(df_raw)

                    df = df_raw.copy()
                    df.columns = df.iloc[linha_header]
                    df = df[(linha_header + 1):]
                    df = df.reset_index(drop=True)

                    return df

                except Exception as e:
                    st.error(f"Erro ao abrir arquivo HTML/XLS: {e}")
                    st.stop()

        else:
            st.error("Formato não suportado.")
            st.stop()

    except Exception as e:
        st.error(f"Erro ao processar arquivo: {e}")
        st.stop()

# =========================
# MENU
# =========================

pagina = st.sidebar.selectbox(
    "Selecione a página",
    [
        "Upload Perfil",
        "Upload Saldo",
        "Upload Parâmetros",
        "Upload Previsão Produção"
    ]
)

st.title("Base de Dados PCP")
st.caption(f"Última atualização: {data_hora_brasil()}")

# =========================
# PERFIL
# =========================

if pagina == "Upload Perfil":

    st.header("Upload Perfil de Itens")

    file = st.file_uploader(
        "Selecione o arquivo Perfil",
        type=["xlsx", "xls", "csv"]
    )

    if file is not None:

        df = carregar_dataframe(file)

        st.success("Arquivo carregado com sucesso!")
        st.write("Pré-visualização:")
        st.dataframe(df.head())

# =========================
# SALDO
# =========================

elif pagina == "Upload Saldo":

    st.header("Upload Saldo de Estoque")

    file = st.file_uploader(
        "Selecione o arquivo Saldo",
        type=["xlsx", "xls", "csv"]
    )

    if file is not None:

        df = carregar_dataframe(file)

        st.success("Arquivo carregado com sucesso!")
        st.write("Pré-visualização:")
        st.dataframe(df.head())

# =========================
# PARÂMETROS
# =========================

elif pagina == "Upload Parâmetros":

    st.header("Upload Parâmetros")

    file = st.file_uploader(
        "Selecione o arquivo de Parâmetros",
        type=["xlsx", "xls", "csv"]
    )

    if file is not None:

        df = carregar_dataframe(file)

        st.success("Arquivo carregado com sucesso!")
        st.write("Pré-visualização:")
        st.dataframe(df.head())

# =========================
# PREVISÃO PRODUÇÃO
# =========================

elif pagina == "Upload Previsão Produção":

    st.header("Upload Previsão Produção")

    file = st.file_uploader(
        "Selecione o arquivo de Previsão",
        type=["xlsx", "xls", "csv"]
    )

    if file is not None:

        df = carregar_dataframe(file)

        st.success("Arquivo carregado com sucesso!")
        st.write("Pré-visualização:")
        st.dataframe(df.head())
