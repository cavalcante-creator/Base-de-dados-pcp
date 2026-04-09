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
# ABAS
# ==========================================================
abas = st.tabs([
    "📦 Saldo",
    "📊 Perfil",
    "📄 Ordens",
    "🗓️ Previsão",
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
                linha = linha.strip()
                if not linha:
                    continue

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
                        if re.search(r'ALMOXARIFADO\s*:\s*3\b', linha.upper()):
                            dados[codigo_atual]["Saldo Almox 3"] += valor

            df = pd.DataFrame(dados.values())
            df["Data Processamento"] = agora().strftime("%d/%m/%Y")
            df["Hora Processamento"] = agora().strftime("%H:%M:%S")

            df.to_csv("saldo.csv", index=False)

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

            regex = re.compile(
                r'(DD|DC|DP).*?(\d{2}/\d{2}/\d{4}).*?(-?[\d,.]+)'
            )

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

            df.to_csv("perfil.csv", index=False)

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

        df.to_csv("ordens.csv", index=False)

        st.success("Ordens carregadas!")
        st.dataframe(df, use_container_width=True)

# ==========================================================
# PREVISÃO (INTELIGENTE)
# ==========================================================
with abas[3]:
    st.title("🗓️ Previsão")

    file = st.file_uploader("Excel Previsão", type=["xlsx"])

    if file:

        df_raw = pd.read_excel(file, header=None)

        linha_header = None
        for i in range(len(df_raw)):
            if "COD" in df_raw.iloc[i].astype(str).str.upper().values:
                linha_header = i
                break

        if linha_header is None:
            st.error("❌ Não foi possível localizar a linha de cabeçalho.")
            st.stop()

        df = pd.read_excel(file, header=linha_header)

        df.columns = df.columns.astype(str).str.upper().str.strip()

        col_cod = [c for c in df.columns if "COD" in c][0]
        col_prod = [c for c in df.columns if "PROD" in c][0]

        df = df[[col_cod, col_prod]]
        df.columns = ["COD", "PRODUTO"]

        df["Data Processamento"] = agora().strftime("%d/%m/%Y")
        df["Hora Processamento"] = agora().strftime("%H:%M:%S")

        df.to_csv("previsao.csv", index=False)

        st.success("Previsão carregada!")
        st.dataframe(df, use_container_width=True)

# ==========================================================
# BASE DE DADOS
# ==========================================================
with abas[4]:
    st.title("📋 Base de Dados (Último Upload)")

    if st.button("🗑️ Limpar Base de Dados"):
        arquivos_para_remover = [
            "saldo.csv",
            "perfil.csv",
            "ordens.csv",
            "previsao.csv"
        ]

        removidos = []

        for arq in arquivos_para_remover:
            if os.path.exists(arq):
                os.remove(arq)
                removidos.append(arq)

        if removidos:
            st.success(f"Arquivos removidos: {', '.join(removidos)}")
        else:
            st.warning("Nenhum arquivo encontrado para remover.")

    arquivos = {
        "Saldo": ("saldo.csv", "Codigo"),
        "Perfil": ("perfil.csv", "Item"),
        "Ordens": ("ordens.csv", None),
        "Previsão": ("previsao.csv", "COD")
    }

    for nome, (arquivo, chave) in arquivos.items():
        st.subheader(nome)

        if os.path.exists(arquivo):
            df = pd.read_csv(arquivo)

            if "Data Processamento" in df.columns:
                df = df.sort_values(
                    by=["Data Processamento", "Hora Processamento"],
                    ascending=False
                )

            if chave and chave in df.columns:
                df = df.drop_duplicates(subset=[chave], keep="first")

            st.dataframe(df, use_container_width=True)

            st.download_button(
                f"📥 Baixar {nome}",
                df.to_csv(index=False).encode("utf-8"),
                file_name=f"{nome}_limpo.csv",
                mime="text/csv"
            )
        else:
            st.warning(f"{nome} ainda não carregado.")

# ==========================================================
# ANALISE PCP
# ==========================================================
with abas[5]:
    st.title("📊 Análise PCP")

    try:
        saldo = pd.read_csv("saldo.csv")
        perfil = pd.read_csv("perfil.csv")
        previsao = pd.read_csv("previsao.csv")
    except:
        st.warning("⚠️ Faça upload dos dados primeiro.")
        st.stop()

    saldo = saldo.sort_values(by=["Data Processamento", "Hora Processamento"], ascending=False)\
        .drop_duplicates(subset=["Codigo"])

    perfil = perfil.sort_values(by=["Data Processamento", "Hora Processamento"], ascending=False)

    previsao = previsao.sort_values(by=["Data Processamento", "Hora Processamento"], ascending=False)\
        .drop_duplicates(subset=["COD"])

    base = previsao[["COD", "PRODUTO"]].copy()
    base.columns = ["Codigo", "Descricao"]

    saldo_base = saldo[["Codigo", "Saldo Total", "Saldo Almox 3"]]

    perfil["Quantidade"] = (
        perfil["Quantidade"]
        .astype(str)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
        .astype(float)
    )

    perfil["Data Fim"] = pd.to_datetime(perfil["Data Fim"], dayfirst=True)

    perfil["Referencia"] = (
        perfil["Data Fim"].dt.isocalendar().week.astype(str).str.zfill(2)
        + "." +
        perfil["Data Fim"].dt.year.astype(str)
    )

    semana = datetime.now().isocalendar()[1]
    ano = datetime.now().year
    ref = str(semana).zfill(2) + "." + str(ano)

    dc = perfil[perfil["Tipo"] == "DC"].groupby("Item")["Quantidade"].sum().reset_index()
    dc.columns = ["Codigo", "Demanda Pedido"]

    dp = perfil[perfil["Tipo"] == "DP"].groupby("Item")["Quantidade"].sum().reset_index()
    dp.columns = ["Codigo", "Demanda DP"]

    dp_sem = perfil[
        (perfil["Tipo"] == "DP") &
        (perfil["Referencia"] == ref)
    ].groupby("Item")["Quantidade"].sum().reset_index()
    dp_sem.columns = ["Codigo", "DP Semana Atual"]

    df = base.merge(saldo_base, on="Codigo", how="left")
    df = df.merge(dc, on="Codigo", how="left")
    df = df.merge(dp, on="Codigo", how="left")
    df = df.merge(dp_sem, on="Codigo", how="left")

    df = df.fillna(0)

    df["Saldo vs Demanda"] = df["Saldo Almox 3"] - df["Demanda Pedido"]

    def status(row):
        if row["Saldo vs Demanda"] < 0:
            return "🔴 FALTA"
        elif row["Demanda Pedido"] >= row["Saldo Almox 3"] * 0.5:
            return "🟡 RISCO"
        else:
            return "🟢 OK"

    df["Status"] = df.apply(status, axis=1)

    st.dataframe(df, use_container_width=True)
