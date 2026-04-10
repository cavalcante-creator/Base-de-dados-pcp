import streamlit as st
import pandas as pd
from datetime import datetime
import pytz

fuso = pytz.timezone("America/Sao_Paulo")

def agora():
    return datetime.now(fuso)

if "previsao_df" not in st.session_state:
    st.session_state["previsao_df"] = pd.DataFrame()

if "saldo_df" not in st.session_state:
    st.session_state["saldo_df"] = pd.DataFrame()

if "perfil_df" not in st.session_state:
    st.session_state["perfil_df"] = pd.DataFrame()

previsao = st.session_state["previsao_df"]
saldo = st.session_state["saldo_df"]
perfil = st.session_state["perfil_df"]

st.set_page_config(page_title="Dashboard PCP", layout="wide")

st.title("Dashboard PCP")

st.info("O dashboard sempre utiliza automaticamente o último processamento disponível na base de dados.")

if previsao.empty:
    st.error("Nenhuma previsão encontrada. Faça o upload e processamento da previsão.")
    st.stop()

if saldo.empty:
    st.error("Nenhum saldo encontrado. Faça o upload e processamento do saldo.")
    st.stop()

if perfil.empty:
    st.error("Nenhum perfil encontrado. Faça o upload e processamento do perfil.")
    st.stop()

if "Data Processamento" in previsao.columns:
    previsao["Data Processamento"] = pd.to_datetime(
        previsao["Data Processamento"],
        errors="coerce"
    )

    previsao = previsao.sort_values(
        by="Data Processamento",
        ascending=False
    ).drop_duplicates(subset=["COD"])

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

if "filtro_status_dashboard" not in st.session_state:
    st.session_state["filtro_status_dashboard"] = "TODOS"

st.subheader("Resumo Geral")

col1, col2, col3, col4 = st.columns(4)

with col1:
    if st.button(f"Todos\n{len(df)}"):
        st.session_state["filtro_status_dashboard"] = "TODOS"

with col2:
    if st.button(f"Falta\n{int((df['Status'] == 'FALTA').sum())}"):
        st.session_state["filtro_status_dashboard"] = "FALTA"

with col3:
    if st.button(f"Risco\n{int((df['Status'] == 'RISCO').sum())}"):
        st.session_state["filtro_status_dashboard"] = "RISCO"

with col4:
    if st.button(f"OK\n{int((df['Status'] == 'OK').sum())}"):
        st.session_state["filtro_status_dashboard"] = "OK"

status_ativo = st.session_state["filtro_status_dashboard"]

st.markdown(f"### Filtro Atual: {status_ativo}")

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

if status_ativo != "TODOS":
    df_filtrado = df_filtrado[df_filtrado["Status"] == status_ativo]

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

st.subheader("Tabela de Análise")

styled_df = df_filtrado.style.map(
    colorir_status,
    subset=["Status"]
).map(
    colorir_saldo,
    subset=["Saldo Total", "Saldo Almox 3", "Saldo vs Demanda"]
)

st.dataframe(
    styled_df,
    use_container_width=True,
    height=600
)
