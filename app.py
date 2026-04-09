import streamlit as st
import pdfplumber
import pandas as pd
import re
from datetime import datetime
import os
from io import StringIO
import pytz

st.set_page_config(page_title="PCP Produção", layout="wide")

fuso = pytz.timezone("America/Sao_Paulo")

def agora():
    return datetime.now(fuso)

# ==========================================================
# 🔥 LEITURA ROBUSTA CSV
# ==========================================================
def ler_csv_seguro(file):
    try:
        return pd.read_csv(file, sep=None, engine="python", encoding="latin-1", on_bad_lines="skip")
    except:
        try:
            return pd.read_csv(file, sep=";", encoding="latin-1", on_bad_lines="skip")
        except:
            try:
                return pd.read_csv(file, sep=",", encoding="latin-1", on_bad_lines="skip")
            except:
                return pd.DataFrame()

# ==========================================================
# SALVAR CSV
# ==========================================================
def salvar_csv(df, nome):
    if os.path.exists(nome):
        antigo = ler_csv_seguro(nome)
        df = pd.concat([antigo, df], ignore_index=True)
    df.to_csv(nome, index=False)

# ==========================================================
# LIMPAR BASE
# ==========================================================
def limpar_base():
    arquivos = ["saldo.csv","perfil.csv","ordens.csv","previsao.csv","parametros.csv"]
    for arq in arquivos:
        if os.path.exists(arq):
            os.remove(arq)

# ==========================================================
# ABAS
# ==========================================================
abas = st.tabs([
    "📦 Saldo",
    "📊 Perfil",
    "📄 Ordens",
    "📅 Previsão",
    "⚙️ Parâmetros",
    "📋 Base de Dados",
    "📊 Dashboard PCP"
])

# ==========================================================
# BASE DE DADOS (VISUALIZAÇÃO)
# ==========================================================
with abas[5]:
    st.title("📋 Base de Dados")

    if st.button("🗑 Limpar Base"):
        limpar_base()
        st.success("Base limpa!")

    for arq in ["saldo.csv","perfil.csv","ordens.csv","previsao.csv","parametros.csv"]:
        st.subheader(arq)

        if os.path.exists(arq):
            df = ler_csv_seguro(arq)
            st.dataframe(df, use_container_width=True)
        else:
            st.warning("Sem dados")

# ==========================================================
# DASHBOARD COMPLETO (IGUAL POWER BI)
# ==========================================================
with abas[6]:
    st.title("📊 Dashboard PCP")

    arquivos = ["saldo.csv","perfil.csv","previsao.csv"]

    for arq in arquivos:
        if not os.path.exists(arq):
            st.warning("Faça upload dos dados primeiro")
            st.stop()

    saldo = ler_csv_seguro("saldo.csv")
    perfil = ler_csv_seguro("perfil.csv")
    previsao = ler_csv_seguro("previsao.csv")

    # 🔥 FILTRO POR DATA (IGUAL POWER BI)
    datas = sorted(saldo["Data Processamento"].dropna().unique())
    data_sel = st.selectbox("📅 Data", datas)

    saldo = saldo[saldo["Data Processamento"] == data_sel].drop_duplicates("Codigo")
    perfil = perfil[perfil["Data Processamento"] == data_sel]
    previsao = previsao[previsao["Data Processamento"] == data_sel].drop_duplicates("COD")

    base = previsao.rename(columns={"COD":"Codigo","PRODUTO":"Descricao"})
    saldo_base = saldo[["Codigo","Saldo Total","Saldo Almox 3"]]

    # 🔥 TRATAR NUMERO
    perfil["Quantidade"] = (
        perfil["Quantidade"].astype(str)
        .str.replace(".", "")
        .str.replace(",", ".")
        .astype(float)
    )

    # 🔥 DEMANDA DC
    dc = perfil[perfil["Tipo"]=="DC"].groupby("Item")["Quantidade"].sum().reset_index()
    dc.columns = ["Codigo","Demanda"]

    # 🔥 JOIN
    df = base.merge(saldo_base, on="Codigo", how="left")
    df = df.merge(dc, on="Codigo", how="left")

    df = df.fillna(0)

    # 🔥 MÉTRICAS
    df["Saldo vs Demanda"] = df["Saldo Almox 3"] - df["Demanda"]

    def status(x):
        if x < 0:
            return "🔴 FALTA"
        elif x < 50:
            return "🟡 RISCO"
        else:
            return "🟢 OK"

    df["Status"] = df["Saldo vs Demanda"].apply(status)

    # ======================================================
    # 🎯 CARDS INTERATIVOS
    # ======================================================
    if "filtro" not in st.session_state:
        st.session_state.filtro = "TODOS"

    col1, col2, col3, col4 = st.columns(4)

    if col1.button(f"🔴 FALTA ({(df['Status']=='🔴 FALTA').sum()})"):
        st.session_state.filtro = "🔴 FALTA"

    if col2.button(f"🟡 RISCO ({(df['Status']=='🟡 RISCO').sum()})"):
        st.session_state.filtro = "🟡 RISCO"

    if col3.button(f"🟢 OK ({(df['Status']=='🟢 OK').sum()})"):
        st.session_state.filtro = "🟢 OK"

    if col4.button("🔄 LIMPAR"):
        st.session_state.filtro = "TODOS"

    # 🔥 FILTRO
    if st.session_state.filtro != "TODOS":
        df = df[df["Status"] == st.session_state.filtro]

    # ======================================================
    # 🎨 TABELA FINAL (IGUAL POWER BI)
    # ======================================================
    st.dataframe(df, use_container_width=True)
