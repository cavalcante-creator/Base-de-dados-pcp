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
                    # ==========================================================
# DASHBOARD PCP (IGUAL POWER BI)
# ==========================================================
if pagina == "📊 Dashboard PCP":
    st.title("📊 Monitoramento PCP")
    st.caption(f"📅 {agora().strftime('%d/%m/%Y')} | ⏰ {agora().strftime('%H:%M:%S')}")

    # =========================
    # CARREGAR DADOS
    # =========================
    try:
        saldo = pd.read_csv("saldo.csv")
        perfil = pd.read_csv("perfil.csv")
    except:
        st.warning("⚠️ Carregue os dados primeiro nas outras abas.")
        st.stop()

    # =========================
    # TRATAMENTO
    # =========================
    perfil["Quantidade"] = (
        perfil["Quantidade"]
        .astype(str)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
        .astype(float)
    )

    perfil["Data Fim"] = pd.to_datetime(perfil["Data Fim"], dayfirst=True, errors="coerce")

    # REFERÊNCIA (SEMANA.ANO)
    perfil["Referência"] = (
        perfil["Data Fim"].dt.isocalendar().week.astype(str).str.zfill(2)
        + "." +
        perfil["Data Fim"].dt.year.astype(str)
    )

    # =========================
    # MEDIDAS (IGUAL DAX)
    # =========================

    # DEMANDA DC
    demanda_dc = perfil[perfil["Tipo"] == "DC"] \
        .groupby("Item")["Quantidade"].sum().reset_index()

    # DEMANDA DP TOTAL
    demanda_dp = perfil[perfil["Tipo"] == "DP"] \
        .groupby("Item")["Quantidade"].sum().reset_index()

    # DEMANDA DP SEMANA ATUAL
    hoje = datetime.now()
    semana = hoje.isocalendar()[1]
    ano = hoje.year
    ref_atual = f"{semana:02d}.{ano}"

    demanda_dp_semana = perfil[
        (perfil["Tipo"] == "DP") &
        (perfil["Referência"] == ref_atual)
    ].groupby("Item")["Quantidade"].sum().reset_index()

    # SALDO ALMOX 3
    if "Saldo Almox 3" in saldo.columns:
        saldo_almox3 = saldo.groupby("Codigo")["Saldo Almox 3"].sum().reset_index()
    else:
        st.error("⚠️ Coluna 'Saldo Almox 3' não encontrada no saldo.")
        st.stop()

    # =========================
    # MERGE GERAL
    # =========================
    df = saldo_almox3.merge(demanda_dc, left_on="Codigo", right_on="Item", how="left")
    df = df.merge(demanda_dp, left_on="Codigo", right_on="Item", how="left", suffixes=("", "_DP"))
    df = df.merge(demanda_dp_semana, left_on="Codigo", right_on="Item", how="left", suffixes=("", "_DP_SEM"))

    # LIMPEZA
    df["Quantidade"] = df["Quantidade"].fillna(0)
    df["Quantidade_DP"] = df["Quantidade_DP"].fillna(0)
    df["Quantidade_DP_SEM"] = df["Quantidade_DP_SEM"].fillna(0)

    # =========================
    # CÁLCULOS
    # =========================

    # SALDO VS DEMANDA
    df["Saldo vs Demanda"] = df["Saldo Almox 3"] - df["Quantidade"]

    # SUGESTÃO PRODUÇÃO (SEM PARAMETROS POR ENQUANTO)
    df["Sugestão Produção"] = df["Quantidade"] - df["Saldo Almox 3"]
    df["Sugestão Produção"] = df["Sugestão Produção"].apply(lambda x: max(x, 0))

    # STATUS
    def definir_status(row):
        if row["Saldo vs Demanda"] < 0:
            return "⛔ FALTA"
        elif row["Quantidade"] >= row["Saldo Almox 3"] * 0.5:
            return "⚠️ RISCO"
        else:
            return "✅ OK"

    df["Status"] = df.apply(definir_status, axis=1)

    # =========================
    # CARDS
    # =========================
    col1, col2, col3 = st.columns(3)

    col1.metric("⛔ FALTA", len(df[df["Status"].str.contains("FALTA")]))
    col2.metric("⚠️ RISCO", len(df[df["Status"].str.contains("RISCO")]))
    col3.metric("✅ OK", len(df[df["Status"].str.contains("OK")]))

    st.divider()

    # =========================
    # CORES
    # =========================
    def cor_status(val):
        if "OK" in val:
            return "background-color: #c6efce; color: #006100"
        elif "RISCO" in val:
            return "background-color: #ffeb9c; color: #9c5700"
        elif "FALTA" in val:
            return "background-color: #ffc7ce; color: #9c0006"

    # =========================
    # TABELA PRINCIPAL
    # =========================
    st.subheader("📋 Consumo / Situação")

    tabela = df[[
        "Codigo",
        "Saldo Almox 3",
        "Quantidade",
        "Quantidade_DP",
        "Quantidade_DP_SEM",
        "Sugestão Produção",
        "Status"
    ]]

    tabela.columns = [
        "Código",
        "Saldo Almox 3",
        "Demanda DC",
        "Demanda DP",
        "DP Semana Atual",
        "Sugestão Produção",
        "Status"
    ]

    st.dataframe(
        tabela.style.applymap(cor_status, subset=["Status"]),
        use_container_width=True
    )

    # =========================
    # GRÁFICO
    # =========================
    st.subheader("📊 Saldo x Demanda")

    grafico = df.set_index("Codigo")[["Saldo Almox 3", "Quantidade"]]
    st.bar_chart(grafico)

    # =========================
    # TOP PRODUÇÃO
    # =========================
    st.subheader("🏭 Sugestão de Produção")

    top = df.sort_values(by="Sugestão Produção", ascending=False).head(10)

    st.dataframe(
        top[["Codigo", "Sugestão Produção", "Status"]],
        use_container_width=True
    )
