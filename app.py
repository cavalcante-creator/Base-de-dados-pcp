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
# SALVAR CSV COM HISTÓRICO
# ==========================================================
def salvar_csv(df, nome):
    if os.path.exists(nome):
        antigo = pd.read_csv(nome)
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
    "📊 Análise PCP"
])

# ==========================================================
# SALDO
# ==========================================================
with abas[0]:
    st.title("📦 Saldo Produção")

    file = st.file_uploader("PDF Saldo", type=["pdf"])

    if file:
        with open("saldo_temp.pdf", "wb") as f:
            f.write(file.read())

        if st.button("Processar Saldo"):
            linhas = []

            with pdfplumber.open("saldo_temp.pdf") as pdf:
                for p in pdf.pages:
                    texto = p.extract_text()
                    if texto:
                        linhas.extend(texto.split("\n"))

            dados = {}
            codigo_atual = None

            for linha in linhas:
                codigo_match = re.search(r'\b([A-Z]{1,3}\d{3,5})\b', linha)

                if codigo_match:
                    codigo_atual = codigo_match.group(1)
                    if codigo_atual not in dados:
                        dados[codigo_atual] = {
                            "Codigo": codigo_atual,
                            "Saldo Total": 0,
                            "Saldo Almox 3": 0
                        }
                    continue

                if "ALMOXARIFADO" in linha.upper() and codigo_atual:
                    nums = re.findall(r'[\d\.]+\,\d+', linha)
                    if nums:
                        valor = float(nums[-1].replace(".", "").replace(",", "."))
                        dados[codigo_atual]["Saldo Total"] += valor
                        dados[codigo_atual]["Saldo Almox 3"] = valor

            df = pd.DataFrame(dados.values())
            df["Data Processamento"] = agora().strftime("%d/%m/%Y")
            df["Hora Processamento"] = agora().strftime("%H:%M:%S")

            salvar_csv(df, "saldo.csv")

            st.success("Saldo processado!")
            st.dataframe(df, use_container_width=True)

# ==========================================================
# PERFIL
# ==========================================================
with abas[1]:
    st.title("📊 Perfil Produção")

    file = st.file_uploader("PDF Perfil", type=["pdf"])

    if file:
        with open("perfil_temp.pdf", "wb") as f:
            f.write(file.read())

        if st.button("Processar Perfil"):

            movimentacoes = []
            codigo_item = ""

            regex = re.compile(r'(DD|DC|DP).*?(\d{2}/\d{2}/\d{4}).*?(-?[\d,.]+)')

            with pdfplumber.open("perfil_temp.pdf") as pdf:
                for p in pdf.pages:
                    texto = p.extract_text()
                    if texto:
                        for linha in texto.split("\n"):
                            item_match = re.search(r'Item:\s*(\S+)', linha)
                            if item_match:
                                codigo_item = item_match.group(1)

                            mov = regex.search(linha)
                            if mov:
                                movimentacoes.append({
                                    "Item": codigo_item,
                                    "Tipo": mov.group(1),
                                    "Data Fim": mov.group(2),
                                    "Quantidade": mov.group(3)
                                })

            df = pd.DataFrame(movimentacoes)
            df["Data Processamento"] = agora().strftime("%d/%m/%Y")
            df["Hora Processamento"] = agora().strftime("%H:%M:%S")

            salvar_csv(df, "perfil.csv")

            st.success("Perfil processado!")
            st.dataframe(df, use_container_width=True)

# ==========================================================
# ORDENS
# ==========================================================
with abas[2]:
    st.title("📄 Ordens")

    file = st.file_uploader("CSV Ordens", type=["csv"])

    if file:
        conteudo = file.read().decode("utf-8", errors="ignore")
        df = pd.read_csv(StringIO(conteudo), sep=None, engine="python")

        df["Data Processamento"] = agora().strftime("%d/%m/%Y")
        df["Hora Processamento"] = agora().strftime("%H:%M:%S")

        salvar_csv(df, "ordens.csv")

        st.success("Ordens carregadas!")
        st.dataframe(df, use_container_width=True)

# ==========================================================
# PREVISÃO
# ==========================================================
with abas[3]:
    st.title("📅 Previsão")

    file = st.file_uploader("Excel Previsão", type=["xlsx"])

    if file:
        df = pd.read_excel(file)
        df.columns = df.columns.str.upper()

        col_cod = [c for c in df.columns if "COD" in c][0]
        col_prod = [c for c in df.columns if "PROD" in c][0]

        df = df[[col_cod, col_prod]]
        df.columns = ["COD", "PRODUTO"]

        df["Data Processamento"] = agora().strftime("%d/%m/%Y")
        df["Hora Processamento"] = agora().strftime("%H:%M:%S")

        salvar_csv(df, "previsao.csv")

        st.success("Previsão carregada!")
        st.dataframe(df, use_container_width=True)

# ==========================================================
# PARÂMETROS
# ==========================================================
with abas[4]:
    st.title("⚙️ Parâmetros")

    file = st.file_uploader("Excel Parâmetros", type=["xlsx"])

    if file:
        df = pd.read_excel(file)
        df.columns = df.columns.str.upper()

        col_cod = [c for c in df.columns if "COD" in c][0]
        col_seg = [c for c in df.columns if "SEG" in c or "ESTQ" in c][0]

        df = df[[col_cod, col_seg]]
        df.columns = ["COD", "ESTQ SEG"]

        df["Data Processamento"] = agora().strftime("%d/%m/%Y")
        df["Hora Processamento"] = agora().strftime("%H:%M:%S")

        salvar_csv(df, "parametros.csv")

        st.success("Parâmetros carregados!")
        st.dataframe(df, use_container_width=True)

# ==========================================================
# BASE DE DADOS
# ==========================================================
with abas[5]:
    st.title("📋 Base de Dados")

    if st.button("🗑 Limpar Base de Dados"):
        limpar_base()
        st.success("Base limpa!")

    for arq in ["saldo.csv","perfil.csv","ordens.csv","previsao.csv","parametros.csv"]:
        st.subheader(arq)

        if os.path.exists(arq):
            df = pd.read_csv(arq)
            st.dataframe(df, use_container_width=True)
        else:
            st.warning("Arquivo não encontrado")

# ==========================================================
# ANALISE PCP
# ==========================================================
with abas[6]:
    st.title("📊 Análise PCP")

    try:
        saldo = pd.read_csv("saldo.csv")
        perfil = pd.read_csv("perfil.csv")
        previsao = pd.read_csv("previsao.csv")
        parametros = pd.read_csv("parametros.csv")
    except:
        st.warning("Faça upload dos dados primeiro")
        st.stop()

    datas = sorted(saldo["Data Processamento"].dropna().unique())
    data_sel = st.selectbox("📅 Data", datas)

    saldo = saldo[saldo["Data Processamento"] == data_sel]\
        .sort_values("Hora Processamento", ascending=False)\
        .drop_duplicates("Codigo")

    perfil = perfil[perfil["Data Processamento"] == data_sel]

    previsao = previsao[previsao["Data Processamento"] == data_sel]\
        .sort_values("Hora Processamento", ascending=False)\
        .drop_duplicates("COD")

    parametros = parametros[parametros["Data Processamento"] == data_sel]\
        .sort_values("Hora Processamento", ascending=False)\
        .drop_duplicates("COD")

    base = previsao.rename(columns={"COD": "Codigo", "PRODUTO": "Descricao"})
    saldo_base = saldo[["Codigo","Saldo Total","Saldo Almox 3"]]
    parametros = parametros.rename(columns={"COD": "Codigo"})

    perfil["Quantidade"] = (
        perfil["Quantidade"].astype(str)
        .str.replace(".", "")
        .str.replace(",", ".")
        .astype(float)
    )

    dc = perfil[perfil["Tipo"] == "DC"].groupby("Item")["Quantidade"].sum().reset_index()
    dc.columns = ["Codigo","Demanda"]

    df = base.merge(saldo_base, on="Codigo", how="left")
    df = df.merge(dc, on="Codigo", how="left")
    df = df.merge(parametros, on="Codigo", how="left")

    df = df.fillna(0)

    df["Saldo vs Demanda"] = df["Saldo Almox 3"] - df["Demanda"]

    def status(x):
        if x < 0:
            return "🔴 FALTA"
        elif x < 50:
            return "🟡 RISCO"
        else:
            return "🟢 OK"

    df["Status"] = df["Saldo vs Demanda"].apply(status)

    # FILTRO POR CARD
    if "filtro_status" not in st.session_state:
        st.session_state.filtro_status = "TODOS"

    col1, col2, col3, col4 = st.columns(4)

    if col1.button(f"🔴 FALTA ({(df['Status']=='🔴 FALTA').sum()})"):
        st.session_state.filtro_status = "🔴 FALTA"

    if col2.button(f"🟡 RISCO ({(df['Status']=='🟡 RISCO').sum()})"):
        st.session_state.filtro_status = "🟡 RISCO"

    if col3.button(f"🟢 OK ({(df['Status']=='🟢 OK').sum()})"):
        st.session_state.filtro_status = "🟢 OK"

    if col4.button("🔄 LIMPAR FILTRO"):
        st.session_state.filtro_status = "TODOS"

    if st.session_state.filtro_status != "TODOS":
        df = df[df["Status"] == st.session_state.filtro_status]

    st.divider()
    st.dataframe(df, use_container_width=True)
