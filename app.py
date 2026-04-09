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
# 🔥 LEITURA CSV ROBUSTA (RESOLVE TODOS ERROS)
# ==========================================================
def ler_csv_seguro(file):
    try:
        return pd.read_csv(file, sep=None, engine="python", encoding="latin-1")
    except:
        try:
            return pd.read_csv(file, sep=";", encoding="latin-1")
        except:
            try:
                return pd.read_csv(file, sep=",", encoding="latin-1")
            except:
                st.error("❌ Erro ao ler CSV")
                st.stop()

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
    "📋 Base",
    "📊 Dashboard"
])

# ==========================================================
# SALDO
# ==========================================================
with abas[0]:
    st.title("📦 Saldo")

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
    st.title("📊 Perfil")

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
                                    "Quantidade": mov.group(3)
                                })

            df = pd.DataFrame(movimentacoes)
            df["Data Processamento"] = agora().strftime("%d/%m/%Y")

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
        conteudo = file.read().decode("latin-1", errors="ignore")
        df = ler_csv_seguro(StringIO(conteudo))

        df["Data Processamento"] = agora().strftime("%d/%m/%Y")

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
        df.columns = df.columns.astype(str).str.upper()

        col_cod = [c for c in df.columns if "COD" in c][0]
        col_prod = [c for c in df.columns if "PROD" in c][0]

        df = df[[col_cod, col_prod]]
        df.columns = ["COD", "PRODUTO"]

        df["Data Processamento"] = agora().strftime("%d/%m/%Y")

        salvar_csv(df, "previsao.csv")

        st.success("Previsão carregada!")
        st.dataframe(df, use_container_width=True)

# ==========================================================
# PARÂMETROS
# ==========================================================
with abas[4]:
    st.title("⚙️ Parâmetros")

    file = st.file_uploader("Arquivo Parâmetros", type=None)

    if file:
        try:
            try:
                df = pd.read_excel(file)
            except:
                df = ler_csv_seguro(file)

            df.columns = df.columns.astype(str).str.upper()

            col_cod = [c for c in df.columns if "COD" in c][0]
            col_seg = [c for c in df.columns if "SEG" in c or "ESTO" in c][0]

            df = df[[col_cod, col_seg]]
            df.columns = ["COD", "ESTQ SEG"]

            df["ESTQ SEG"] = pd.to_numeric(df["ESTQ SEG"], errors="coerce").fillna(0)

            df["Data Processamento"] = agora().strftime("%d/%m/%Y")

            salvar_csv(df, "parametros.csv")

            st.success("Parâmetros carregados!")
            st.dataframe(df)

        except Exception as e:
            st.error(f"Erro: {e}")

# ==========================================================
# BASE
# ==========================================================
with abas[5]:
    st.title("📋 Base")

    if st.button("🗑 Limpar Base"):
        limpar_base()
        st.success("Base limpa!")

    for arq in ["saldo.csv","perfil.csv","ordens.csv","previsao.csv","parametros.csv"]:
        st.subheader(arq)

        if os.path.exists(arq):
            df = ler_csv_seguro(arq)
            st.dataframe(df)
        else:
            st.warning("Sem dados")

# ==========================================================
# DASHBOARD
# ==========================================================
with abas[6]:
    st.title("📊 Dashboard PCP")

    if not os.path.exists("saldo.csv"):
        st.warning("Faça upload dos dados")
        st.stop()

    saldo = ler_csv_seguro("saldo.csv")
    perfil = ler_csv_seguro("perfil.csv")
    previsao = ler_csv_seguro("previsao.csv")

    base = previsao.rename(columns={"COD":"Codigo","PRODUTO":"Descricao"})
    saldo = saldo.drop_duplicates("Codigo")

    perfil["Quantidade"] = (
        perfil["Quantidade"].astype(str)
        .str.replace(".", "")
        .str.replace(",", ".")
        .astype(float)
    )

    dc = perfil[perfil["Tipo"]=="DC"].groupby("Item")["Quantidade"].sum().reset_index()
    dc.columns = ["Codigo","Demanda"]

    df = base.merge(saldo, on="Codigo", how="left")
    df = df.merge(dc, on="Codigo", how="left")

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

    st.dataframe(df, use_container_width=True)
