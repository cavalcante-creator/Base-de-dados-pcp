import streamlit as st
import pandas as pd
from datetime import datetime
import pytz

st.set_page_config(page_title="Dashboard PCP", layout="wide")

st.markdown("""
<style>
    .main {
        background: linear-gradient(180deg, #f7fafc 0%, #eef4f7 100%);
    }

    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }

    h1, h2, h3 {
        color: #12344d;
        font-weight: 700;
    }

    div[data-testid="stAlert"] {
        border-radius: 14px;
        border: 1px solid #d7e3ee;
    }

    div.stButton > button {
        background: linear-gradient(90deg, #0f766e, #0ea5a4);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 0.6rem 1.2rem;
        font-weight: 600;
    }

    div.stButton > button:hover {
        background: linear-gradient(90deg, #0b5f59, #0b8f8d);
        color: white;
    }

    div[data-testid="stDownloadButton"] > button {
        background: white;
        color: #12344d;
        border: 1px solid #c9d9e6;
        border-radius: 12px;
        font-weight: 600;
    }

    div[data-testid="stFileUploader"] {
        background: white;
        border: 1px dashed #aac4d6;
        border-radius: 16px;
        padding: 10px;
    }

    div[data-testid="stDataFrame"] {
        background: white;
        border-radius: 16px;
        padding: 8px;
        border: 1px solid #dbe7f0;
        box-shadow: 0 4px 14px rgba(18, 52, 77, 0.05);
    }
</style>
""", unsafe_allow_html=True)

fuso = pytz.timezone("America/Sao_Paulo")

def agora():
    return datetime.now(fuso)

st.title("Dashboard PCP")

if "previsao_df" not in st.session_state:
    st.session_state["previsao_df"] = pd.DataFrame()

if "saldo_df" not in st.session_state:
    st.session_state["saldo_df"] = pd.DataFrame()

if "perfil_df" not in st.session_state:
    st.session_state["perfil_df"] = pd.DataFrame()

previsao = st.session_state["previsao_df"]
saldo = st.session_state["saldo_df"]
perfil = st.session_state["perfil_df"]

if previsao.empty or saldo.empty or perfil.empty:
    st.warning("Faça o processamento dos arquivos antes de acessar o dashboard.")
    st.stop()

base = previsao[["COD", "PRODUTO"]].copy()
base.columns = ["Codigo", "Descricao"]

saldo_base = saldo[["Codigo", "Saldo Total", "Saldo Almox 3"]].copy()

perfil["Quantidade"] = (
    perfil["Quantidade"]
    .astype(str)
    .str.replace(".", "", regex=False)
    .str.replace(",", ".", regex=False)
)

perfil["Quantidade"] = pd.to_numeric(
    perfil["Quantidade"],
    errors="coerce"
).fillna(0)

perfil["Data Fim"] = pd.to_datetime(
    perfil["Data Fim"],
    dayfirst=True,
    errors="coerce"
)

perfil["Referencia"] = (
    perfil["Data Fim"].dt.isocalendar().week.astype(str).str.zfill(2)
    + "."
    + perfil["Data Fim"].dt.year.astype(str)
)

semana = agora().isocalendar()[1]
ano = agora().year
referencia_atual = str(semana).zfill(2) + "." + str(ano)

dc = (
    perfil[perfil["Tipo"] == "DC"]
    .groupby("Item")["Quantidade"]
    .sum()
    .reset_index()
)
dc.columns = ["Codigo", "Demanda Pedido"]

dp = (
    perfil[perfil["Tipo"] == "DP"]
    .groupby("Item")["Quantidade"]
    .sum()
    .reset_index()
)
dp.columns = ["Codigo", "Demanda DP"]

dp_sem = (
    perfil[
        (perfil["Tipo"] == "DP") &
        (perfil["Referencia"] == referencia_atual)
    ]
    .groupby("Item")["Quantidade"]
    .sum()
    .reset_index()
)
dp_sem.columns = ["Codigo", "DP Semana Atual"]

df = base.merge(saldo_base, on="Codigo", how="left")
df = df.merge(dc, on="Codigo", how="left")
df = df.merge(dp, on="Codigo", how="left")
df = df.merge(dp_sem, on="Codigo", how="left")

colunas_numericas = [
    "Saldo Total",
    "Saldo Almox 3",
    "Demanda Pedido",
    "Demanda DP",
    "DP Semana Atual"
]

for coluna in colunas_numericas:
    df[coluna] = pd.to_numeric(df[coluna], errors="coerce").fillna(0)

df["Saldo vs Demanda"] = df["Saldo Almox 3"] - df["Demanda Pedido"]

def definir_status(row):
    if row["Saldo vs Demanda"] < 0:
        return "FALTA"
    elif row["Demanda Pedido"] >= row["Saldo Almox 3"] * 0.5:
        return "RISCO"
    else:
        return "OK"

df["Status"] = df.apply(definir_status, axis=1)

st.subheader("Resumo Geral")

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric("Total de Itens", len(df))

with col2:
    st.metric("Itens em Falta", int((df['Status'] == 'FALTA').sum()))

with col3:
    st.metric("Itens em Risco", int((df['Status'] == 'RISCO').sum()))

with col4:
    st.metric("Itens OK", int((df['Status'] == 'OK').sum()))

col_busca1, col_busca2 = st.columns([2, 1])

with col_busca1:
    busca = st.text_input("Buscar código ou descrição")

with col_busca2:
    ordenar = st.selectbox(
        "Ordenar por",
        [
            "Codigo",
            "Descricao",
            "Saldo Total",
            "Demanda Pedido",
            "Saldo vs Demanda"
        ]
    )

df_filtrado = df.copy()

if busca:
    busca = busca.lower()

    df_filtrado = df_filtrado[
        df_filtrado["Codigo"].astype(str).str.lower().str.contains(busca, na=False) |
        df_filtrado["Descricao"].astype(str).str.lower().str.contains(busca, na=False)
    ]

df_filtrado = df_filtrado.sort_values(by=ordenar)

def colorir_status(valor):
    if valor == "FALTA":
        return "background-color: #f8d7da; color: #842029; font-weight: bold;"
    elif valor == "RISCO":
        return "background-color: #fff3cd; color: #664d03; font-weight: bold;"
    elif valor == "OK":
        return "background-color: #d1e7dd; color: #0f5132; font-weight: bold;"
    return ""

def colorir_saldo(valor):
    try:
        valor = float(valor)
    except:
        return ""

    if valor < 0:
        return "background-color: #f8d7da; color: #842029;"
    elif valor == 0:
        return "background-color: #fff3cd; color: #664d03;"
    else:
        return "background-color: #d1e7dd; color: #0f5132;"

styled_df = df_filtrado.style.map(
    colorir_status,
    subset=["Status"]
).map(
    colorir_saldo,
    subset=["Saldo Total", "Saldo Almox 3", "Saldo vs Demanda"]
)

st.subheader("Tabela de Análise")

st.dataframe(
    styled_df,
    use_container_width=True,
    height=600
)
