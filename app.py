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
        "Upload Parâmetros",
        "Gerar Excel"
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

            arquivo = "saldo.csv"

            if os.path.exists(arquivo):
                df_antigo = pd.read_csv(arquivo)
                df = pd.concat([df_antigo, df], ignore_index=True)

            df.to_csv(arquivo, index=False)

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

            arquivo = "perfil.csv"

            if os.path.exists(arquivo):
                df_antigo = pd.read_csv(arquivo)
                df = pd.concat([df_antigo, df], ignore_index=True)

            df.to_csv(arquivo, index=False)

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

        arquivo = "ordens.csv"

        if os.path.exists(arquivo):
            df_antigo = pd.read_csv(arquivo)
            df = pd.concat([df_antigo, df], ignore_index=True)

        df.to_csv(arquivo, index=False)

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

    st.title("⚙️ Upload Parâmetros")

    file = st.file_uploader(
        "Excel Parâmetros",
        type=["xlsx", "xls"]
    )

    if file:

        try:
            file.seek(0)
            df_raw = pd.read_excel(file, header=None, engine="openpyxl")
            engine_excel = "openpyxl"
        except:
            file.seek(0)
            df_raw = pd.read_excel(file, header=None, engine="xlrd")
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

                st.success("Parâmetros carregados!")
                st.dataframe(df, use_container_width=True)

# ==========================================================
# GERAR EXCEL (COM HISTÓRICO)
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
