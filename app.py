import streamlit as st
import pdfplumber
import pandas as pd
import re
from datetime import datetime
import os
from io import StringIO

st.set_page_config(page_title="PCP Produção", layout="wide")

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
        "Gerar Excel"
    ]
)

# ==========================================================
# SALDO (COMPLETO)
# ==========================================================

if pagina == "Upload Saldo Produção":

    st.title("📦 Extração de Saldo de Produção")

    data_atual = datetime.now()
    st.caption(f"📅 {data_atual.strftime('%d/%m/%Y')} | ⏰ {data_atual.strftime('%H:%M:%S')}")

    uploaded_file = st.file_uploader("Selecione o PDF", type=["pdf"])

    if uploaded_file:

        with open("saldo_temp.pdf", "wb") as f:
            f.write(uploaded_file.read())

        if st.button("Processar Saldo"):

            linhas = []

            with pdfplumber.open("saldo_temp.pdf") as pdf:
                for pagina_pdf in pdf.pages:
                    texto = pagina_pdf.extract_text()
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

                    descricao = re.split(
                        r'\s+\d+[.,]?\d*|\s+UN\b|\s+KG\b|\s+SC\b|\s+CX\b|\s+PCT\b|\s+M2\b|\s+M3\b',
                        descricao
                    )[0].strip()

                    descricao = re.sub(r'\s+', ' ', descricao)
                    descricao = descricao.strip("-").strip()

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

                        numero_almox = almox_match.group(1)
                        coluna_almox = f"Saldo Almox {numero_almox}"

                        if coluna_almox not in dados[codigo_atual]:
                            dados[codigo_atual][coluna_almox] = 0

                        numeros = re.findall(r'[\d\.]+\,\d+', linha)

                        if numeros:
                            saldo = float(numeros[-1].replace(".", "").replace(",", "."))
                            dados[codigo_atual][coluna_almox] += saldo
                            dados[codigo_atual]["Saldo Total"] += saldo

            df = pd.DataFrame(dados.values())

            colunas_almox = sorted(
                [c for c in df.columns if "Saldo Almox" in c],
                key=lambda x: int(re.search(r'(\d+)', x).group())
            )

            df["Data Processamento"] = datetime.now().strftime("%d/%m/%Y")
            df["Hora Processamento"] = datetime.now().strftime("%H:%M:%S")

            df = df[["Codigo", "Descricao"] + colunas_almox + ["Saldo Total", "Data Processamento", "Hora Processamento"]]

            df.to_csv("saldo.csv", index=False)

            st.dataframe(df, use_container_width=True)

# ==========================================================
# PERFIL (COMPLETO)
# ==========================================================

if pagina == "Upload Perfil Produção":

    st.title("📊 Perfil de Produção")

    uploaded_file = st.file_uploader("Selecione o PDF", type=["pdf"])

    if uploaded_file:

        with open("perfil_temp.pdf", "wb") as f:
            f.write(uploaded_file.read())

        if st.button("Processar Perfil"):

            def extrair_numero_sinal(texto):
                if not texto:
                    return ''
                match = re.search(r'-?[\d,.]+', texto)
                if match:
                    return match.group(0).replace('.', '').replace(',', '.')
                return ''

            movimentacoes = []
            codigo_item = ''

            regex_mov = re.compile(
                r'\b(DD|DC|DP|OCP|OCL|OFP(?:\.\d+)?|OFA(?:\.\d+)?|AVR|TIPO\s+DE)\b\s+'
                r'([\w.-]+)?\s*'
                r'((\d{2}/\d{2}/\d{4})\s+)?'
                r'(\d{2}/\d{2}/\d{4})\s+'
                r'(-?[\d,.]+)\s*\S*\s+(-?[\d,.]+)'
            )

            with pdfplumber.open("perfil_temp.pdf") as pdf:
                for page in pdf.pages:
                    texto = page.extract_text()
                    if texto:
                        for linha in texto.split("\n"):

                            match_item = re.search(r'Item:\s*([\w.-]+)', linha)
                            if match_item:
                                codigo_item = match_item.group(1)
                                continue

                            match_mov = regex_mov.search(linha)
                            if match_mov:

                                movimentacoes.append({
                                    "Item": codigo_item,
                                    "Tipo": match_mov.group(1),
                                    "Referência": match_mov.group(2),
                                    "Data Fim": match_mov.group(5),
                                    "Quantidade": extrair_numero_sinal(match_mov.group(6)),
                                    "Estoque Projetado": extrair_numero_sinal(match_mov.group(7))
                                })

            df = pd.DataFrame(movimentacoes)

            df["Data Processamento"] = datetime.now().strftime("%d/%m/%Y")
            df["Hora Processamento"] = datetime.now().strftime("%H:%M:%S")

            df.to_csv("perfil.csv", index=False)

            st.dataframe(df, use_container_width=True)

# ==========================================================
# ORDENS
# ==========================================================

if pagina == "Upload Ordens de Fabricação":

    st.title("📄 Ordens de Fabricação")

    file = st.file_uploader("CSV", type=["csv"])

    if file:

        conteudo = file.read().decode("utf-8", errors="ignore")

        if not conteudo.strip():
            st.error("CSV vazio")
            st.stop()

        try:
            df = pd.read_csv(StringIO(conteudo), sep=";")
            if df.shape[1] == 1:
                df = pd.read_csv(StringIO(conteudo), sep=",")
        except:
            df = pd.read_csv(StringIO(conteudo))

        df["Data Processamento"] = datetime.now().strftime("%d/%m/%Y")
        df["Hora Processamento"] = datetime.now().strftime("%H:%M:%S")

        df.to_csv("ordens.csv", index=False)

        st.dataframe(df, use_container_width=True)

# ==========================================================
# PREVISÃO (INTELIGENTE)
# ==========================================================

if pagina == "Upload Previsão Produção":

    st.title("📅 Previsão Produção")

    file = st.file_uploader("Excel", type=["xlsx"])

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

        df = df.replace("#ERROR!", 0)
        df["PREVISAO"] = pd.to_numeric(df["PREVISAO"], errors="coerce").fillna(0)

        df.to_csv("previsao.csv", index=False)

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

            if st.button(f"Gerar Excel {nome}"):

                nome_excel = f"{nome}.xlsx"
                df.to_excel(nome_excel, index=False)

                with open(nome_excel, "rb") as f:
                    st.download_button(f"📥 Baixar {nome}", f, file_name=nome_excel)