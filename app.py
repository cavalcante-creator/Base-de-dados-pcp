import streamlit as st
import pdfplumber
import pandas as pd
import re
from datetime import datetime
import os
from io import StringIO
import pytz

st.set_page_config(page_title="PCP Produção", layout="wide")

# FUSO BRASIL
fuso = pytz.timezone("America/Sao_Paulo")

def agora():
    return datetime.now(fuso)

# ==========================================================
# MENU
# ==========================================================
pagina = st.sidebar.radio(
    "Menu",
    [
        "Upload Saldo Produção",
        "Upload Perfil Produção",
        "Upload Ordens de Fabricação",
        "Upload Previsão Produção",
        "Gerar Excel",
        "📊 Análise PCP"
    ]
)

# ==========================================================
# SALDO
# ==========================================================
if pagina == "Upload Saldo Produção":
    st.title("📦 Saldo Produção")
    st.caption(f"📅 {agora().strftime('%d/%m/%Y')} | ⏰ {agora().strftime('%H:%M:%S')}")

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

                linha_upper = linha.upper()

                codigo_match = re.search(r'\b([A-Z]{1,3}\d{3,5})\b', linha)

                if codigo_match:
                    codigo_atual = codigo_match.group(1)
                    descricao = linha.split(codigo_atual, 1)[1].strip()
                    descricao = re.split(r'\s+\d+[.,]?\d*|\s+UN\b|\s+KG\b', descricao)[0].strip()

                    if codigo_atual not in dados:
                        dados[codigo_atual] = {
                            "Codigo": codigo_atual,
                            "Descricao": descricao,
                            "Saldo Total": 0
                        }
                    continue

                if "ALMOXARIFADO" in linha_upper and codigo_atual:
                    almox_match = re.search(r'ALMOXARIFADO[:\s]+(\d+)', linha_upper)
                    if almox_match:
                        numero = almox_match.group(1)
                        col = f"Saldo Almox {numero}"

                        if col not in dados[codigo_atual]:
                            dados[codigo_atual][col] = 0

                        nums = re.findall(r'[\d\.]+\,\d+', linha)
                        if nums:
                            valor = float(nums[-1].replace(".", "").replace(",", "."))
                            dados[codigo_atual][col] += valor
                            dados[codigo_atual]["Saldo Total"] += valor

            df = pd.DataFrame(dados.values())
            df["Data Processamento"] = agora().strftime("%d/%m/%Y")
            df["Hora Processamento"] = agora().strftime("%H:%M:%S")

            df.to_csv("saldo.csv", index=False)

            st.success("Saldo processado!")
            st.dataframe(df, use_container_width=True)

# ==========================================================
# PERFIL
# ==========================================================
if pagina == "Upload Perfil Produção":
    st.title("📊 Perfil Produção")
    st.caption(f"📅 {agora().strftime('%d/%m/%Y')} | ⏰ {agora().strftime('%H:%M:%S')}")

    file = st.file_uploader("PDF Perfil", type=["pdf"])

    if file:
        with open("perfil_temp.pdf", "wb") as f:
            f.write(file.read())

        if st.button("Processar Perfil"):

            def extrair_numero(texto):
                m = re.search(r'-?[\d,.]+', str(texto))
                return m.group(0) if m else "0"

            movimentacoes = []
            codigo_item = ""

            regex = re.compile(
                r'(DD|DC|DP|OFP|OFA).*?(\d{2}/\d{2}/\d{4}).*?(-?[\d,.]+).*?(-?[\d,.]+)'
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
                                    "Quantidade": extrair_numero(mov.group(3)),
                                    "Estoque Projetado": extrair_numero(mov.group(4))
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
if pagina == "Upload Ordens de Fabricação":
    st.title("📄 Ordens de Fabricação")

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
# PREVISÃO
# ==========================================================
if pagina == "Upload Previsão Produção":
    st.title("📅 Previsão Produção")

    file = st.file_uploader("Excel Previsão", type=["xlsx"])

    if file:
        df_raw = pd.read_excel(file, header=None)

        linha_header = None
        for i in range(len(df_raw)):
            if "COD" in df_raw.iloc[i].astype(str).str.upper().values:
                linha_header = i
                break

        df = pd.read_excel(file, header=linha_header)
        df.columns = df.columns.astype(str).str.upper()

        col_cod = [c for c in df.columns if "COD" in c][0]
        col_prod = [c for c in df.columns if "PROD" in c][0]
        col_prev = [c for c in df.columns if "PREVIS" in c][0]

        df = df[[col_cod, col_prod, col_prev]]
        df.columns = ["COD", "PRODUTO", "PREVISAO"]

        df["PREVISAO"] = pd.to_numeric(df["PREVISAO"], errors="coerce").fillna(0)

        df["Data Processamento"] = agora().strftime("%d/%m/%Y")
        df["Hora Processamento"] = agora().strftime("%H:%M:%S")

        df.to_csv("previsao.csv", index=False)

        st.success("Previsão carregada!")
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
        "Previsão": "previsao.csv"
    }

    for nome, arquivo in arquivos.items():
        if os.path.exists(arquivo):
            df = pd.read_csv(arquivo)

            st.subheader(nome)
            st.dataframe(df, use_container_width=True)

# ==========================================================
# ANALISE PCP (BASE CORRETA)
# ==========================================================
if pagina == "📊 Análise PCP":
    st.title("📊 Análise PCP")

    try:
        saldo = pd.read_csv("saldo.csv")
        perfil = pd.read_csv("perfil.csv")
        previsao = pd.read_csv("previsao.csv")
    except:
        st.warning("⚠️ Carregue todos os dados primeiro.")
        st.stop()

    # =========================
    # ULTIMO UPLOAD
    # =========================
    saldo = saldo.sort_values(by=["Data Processamento", "Hora Processamento"], ascending=False)\
        .drop_duplicates(subset=["Codigo"])

    perfil = perfil.sort_values(by=["Data Processamento", "Hora Processamento"], ascending=False)

    previsao = previsao.sort_values(by=["Data Processamento", "Hora Processamento"], ascending=False)\
        .drop_duplicates(subset=["COD"])

    # =========================
    # BASE PREVISÃO
    # =========================
    base = previsao[["COD", "PRODUTO"]].copy()
    base.columns = ["Codigo", "Descricao"]

    # =========================
    # SALDO
    # =========================
    saldo_base = saldo[["Codigo", "Saldo Total", "Saldo Almox 3"]]

    # =========================
    # DEMANDA DC
    # =========================
    perfil["Quantidade"] = (
        perfil["Quantidade"]
        .astype(str)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
        .astype(float)
    )

    demanda = perfil[perfil["Tipo"] == "DC"]\
        .groupby("Item")["Quantidade"].sum().reset_index()

    demanda.columns = ["Codigo", "Demanda Pedido"]

    # =========================
    # MERGE FINAL
    # =========================
    df = base.merge(saldo_base, on="Codigo", how="left")
    df = df.merge(demanda, on="Codigo", how="left")

    df = df.fillna(0)

    st.dataframe(df, use_container_width=True)
