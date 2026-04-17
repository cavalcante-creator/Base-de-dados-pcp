import re
from datetime import datetime
from io import BytesIO

import pandas as pd
import pytz
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

st.set_page_config(page_title="Dashboard PCP", layout="wide")

# ========================= CSS =========================
st.markdown("""
<style>
.card {
    border-radius: 14px;
    padding: 14px;
    text-align: center;
    color: white;
    font-weight: 600;
}

.total {background:#1d4ed8;}
.falta {background:#dc2626;}
.risco {background:#f59e0b;}
.ok {background:#16a34a;}

button[kind="secondary"] {
    width: 100%;
    border-radius: 14px;
    height: 80px;
    font-weight: bold;
}
</style>
""", unsafe_allow_html=True)

# ========================= FUNÇÕES =========================
fuso = pytz.timezone("America/Sao_Paulo")

def agora():
    return datetime.now(fuso)

def exportar_excel_formatado(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Dados")
    output.seek(0)
    return output.getvalue()

# ========================= DADOS =========================
try:
    saldo = pd.read_csv("saldo.csv")
    perfil = pd.read_csv("perfil.csv")
    previsao = pd.read_csv("previsao.csv")
except:
    st.warning("Faça upload dos dados primeiro.")
    st.stop()

try:
    parametros = pd.read_csv("parametros.csv")
    parametros = parametros.rename(columns={"COD ITEM": "Codigo", "ESTQ SEG": "Estq Seg"})
    parametros = parametros[["Codigo", "Estq Seg"]].copy()
    parametros["Estq Seg"] = pd.to_numeric(parametros["Estq Seg"], errors="coerce").fillna(0)
    tem_parametros = True
except:
    tem_parametros = False

saldo = saldo.sort_values(by=["Data Processamento","Hora Processamento"], ascending=False).drop_duplicates("Codigo")
previsao = previsao.sort_values(by=["Data Processamento","Hora Processamento"], ascending=False).drop_duplicates("COD")

base = previsao[["COD","PRODUTO"]].copy()
base.columns = ["Codigo","Descricao"]

saldo_base = saldo[["Codigo","Saldo Total","Saldo Almox 3"]]

perfil["Quantidade"] = (
    perfil["Quantidade"].astype(str)
    .str.replace(".","",regex=False)
    .str.replace(",",".",regex=False)
    .astype(float)
)

# ========================= DEMANDAS =========================
dc = perfil[perfil["Tipo"]=="DC"].groupby("Item")["Quantidade"].sum().reset_index()
dc.columns = ["Codigo","Demanda Pedido"]

# 🔥 ORDENS LIBERADAS (MAIS SEGURO)
tipos_op = ["OP","ORDEM","LIBERADA"]

op = perfil[perfil["Tipo"].isin(tipos_op)].groupby("Item")["Quantidade"].sum().reset_index()
op.columns = ["Codigo","Qtde Pendente OP"]

# ========================= MERGE =========================
df = base.merge(saldo_base,on="Codigo",how="left")
df = df.merge(dc,on="Codigo",how="left")
df = df.merge(op,on="Codigo",how="left")

if tem_parametros:
    df = df.merge(parametros, on="Codigo", how="left")
    df["Estq Seg"] = df["Estq Seg"].fillna(0)
else:
    df["Estq Seg"] = 0

df = df.fillna(0)

# ========================= CÁLCULOS =========================
df["Saldo vs Demanda"] = df["Saldo Almox 3"] - df["Demanda Pedido"]

df["Saldo Real"] = (
    df["Saldo Almox 3"]
    - df["Demanda Pedido"]
    - df["Qtde Pendente OP"]
)

df["Abaixo Estq Seg"] = df["Saldo Real"] < df["Estq Seg"]

# ========================= STATUS =========================
def status(row):
    if row["Saldo Real"] < 0:
        return "FALTA"
    if row["Abaixo Estq Seg"]:
        return "RISCO"
    if row["Demanda Pedido"] + row["Qtde Pendente OP"] >= row["Saldo Almox 3"] * 0.5:
        return "RISCO"
    return "OK"

df["Status"] = df.apply(status, axis=1)

# ========================= CARDS (FUNCIONAIS) =========================
if "filtro" not in st.session_state:
    st.session_state.filtro = "TODOS"

c1,c2,c3,c4 = st.columns(4)

if c1.button(f"TOTAL\n{len(df)}"):
    st.session_state.filtro = "TODOS"

if c2.button(f"FALTA\n{(df['Status']=='FALTA').sum()}"):
    st.session_state.filtro = "FALTA"

if c3.button(f"RISCO\n{(df['Status']=='RISCO').sum()}"):
    st.session_state.filtro = "RISCO"

if c4.button(f"OK\n{(df['Status']=='OK').sum()}"):
    st.session_state.filtro = "OK"

# ========================= FILTRO =========================
if st.session_state.filtro == "TODOS":
    df_filtrado = df.copy()
else:
    df_filtrado = df[df["Status"] == st.session_state.filtro]

# ========================= BUSCA =========================
busca = st.text_input("Buscar")

if busca:
    df_filtrado = df_filtrado[
        df_filtrado["Codigo"].astype(str).str.contains(busca, case=False) |
        df_filtrado["Descricao"].astype(str).str.contains(busca, case=False)
    ]

# ========================= COR NA LINHA =========================
def cor(row):
    if row["Status"]=="FALTA":
        return ["background-color:#fecaca"]*len(row)
    if row["Status"]=="RISCO":
        return ["background-color:#fde68a"]*len(row)
    if row["Status"]=="OK":
        return ["background-color:#bbf7d0"]*len(row)
    return [""]*len(row)

st.dataframe(df_filtrado.style.apply(cor, axis=1), use_container_width=True)

# ========================= ALERTA ESTQ SEG =========================
if tem_parametros:
    df_abaixo = df[df["Abaixo Estq Seg"] & (df["Status"] != "FALTA")].copy()
    if not df_abaixo.empty:
        with st.expander(f"⚠️ {len(df_abaixo)} item(ns) abaixo do Estoque de Segurança", expanded=True):
            cols_alerta = ["Codigo", "Descricao", "Saldo Real", "Estq Seg", "Status"]
            cols_alerta = [c for c in cols_alerta if c in df_abaixo.columns]
            df_alerta = df_abaixo[cols_alerta].copy()
            df_alerta["Diferença"] = df_alerta["Saldo Real"] - df_alerta["Estq Seg"]
            st.dataframe(df_alerta.style.apply(cor, axis=1), use_container_width=True)
    else:
        st.success("✅ Todos os itens estão acima do Estoque de Segurança.")
elif not tem_parametros:
    st.info("ℹ️ Importe o arquivo de Parâmetros para comparar com o Estoque de Segurança.")

# ========================= DOWNLOAD =========================
st.download_button("Baixar CSV", df_filtrado.to_csv(index=False), "pcp.csv")
