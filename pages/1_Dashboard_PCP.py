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

# ========================= ORDENS EM ABERTO =========================
tem_ordens = False
ordens_abertas = pd.DataFrame()

try:
    ordens_raw = pd.read_csv("ordens.csv")
    colunas_upper = {c: c.upper().strip() for c in ordens_raw.columns}
    col_cod_ordens = None
    col_qtd_ordens = None
    col_status_ordens = None

    for orig, upper in colunas_upper.items():
        if any(k in upper for k in ["COD", "ITEM", "PRODUTO", "CODIGO", "CÓDIGO"]) and col_cod_ordens is None:
            col_cod_ordens = orig
        if any(k in upper for k in ["QTDE", "QTD", "QUANTIDADE", "SALDO", "SALDO ORDEM"]) and col_qtd_ordens is None:
            col_qtd_ordens = orig
        if any(k in upper for k in ["STATUS", "SITUAÇÃO", "SITUACAO", "SIT"]) and col_status_ordens is None:
            col_status_ordens = orig

    if col_cod_ordens and col_qtd_ordens:
        cols = [col_cod_ordens, col_qtd_ordens] + ([col_status_ordens] if col_status_ordens else [])
        df_ord = ordens_raw[cols].copy()
        df_ord.columns = ["Codigo", "Qtde Ordem"] + (["Status Ordem"] if col_status_ordens else [])

        if "Status Ordem" in df_ord.columns:
            palavras_aberto = ["ABERTO", "ABERTA", "EM ABERTO", "LIBERADA", "LIBERADO", "PENDENTE", "A PRODUZIR"]
            mask_aberto = df_ord["Status Ordem"].astype(str).str.upper().str.strip().isin(palavras_aberto)
            if mask_aberto.sum() > 0:
                df_ord = df_ord[mask_aberto]

        df_ord["Qtde Ordem"] = (
            df_ord["Qtde Ordem"].astype(str)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        df_ord["Qtde Ordem"] = pd.to_numeric(df_ord["Qtde Ordem"], errors="coerce").fillna(0)
        ordens_abertas = df_ord.groupby("Codigo")["Qtde Ordem"].sum().reset_index()
        ordens_abertas.columns = ["Codigo", "Qtde Ordens Abertas"]
        tem_ordens = True
    else:
        st.info("ℹ️ Arquivo de Ordens carregado, mas não foi possível identificar colunas de código/quantidade automaticamente.")
except:
    pass

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

# ========================= DETECTAR ALMOXARIFADOS DISPONÍVEIS =========================
colunas_almox = [c for c in saldo.columns if "SALDO" in c.upper() and "ALMOX" in c.upper()]

if "Saldo Total" in saldo.columns:
    opcoes_almox = ["Saldo Total"] + colunas_almox
else:
    opcoes_almox = colunas_almox

if not opcoes_almox:
    st.error("Nenhuma coluna de saldo encontrada. Reprocesse o PDF de Saldo.")
    st.stop()

for col in opcoes_almox:
    if col not in saldo.columns:
        saldo[col] = 0

saldo_base = saldo[["Codigo"] + opcoes_almox]

perfil["Quantidade"] = (
    perfil["Quantidade"].astype(str)
    .str.replace(".", "", regex=False)
    .str.replace(",", ".", regex=False)
    .astype(float)
)

# ========================= DEMANDAS =========================
dc = perfil[perfil["Tipo"]=="DC"].groupby("Item")["Quantidade"].sum().reset_index()
dc.columns = ["Codigo","Demanda Pedido"]

tipos_op = ["OP","ORDEM","LIBERADA"]
op = perfil[perfil["Tipo"].isin(tipos_op)].groupby("Item")["Quantidade"].sum().reset_index()
op.columns = ["Codigo","Qtde Pendente OP"]

# ========================= MERGE =========================
df = base.merge(saldo_base, on="Codigo", how="left")
df = df.merge(dc, on="Codigo", how="left")
df = df.merge(op, on="Codigo", how="left")

if tem_ordens:
    df = df.merge(ordens_abertas, on="Codigo", how="left")

if tem_parametros:
    df = df.merge(parametros, on="Codigo", how="left")
    df["Estq Seg"] = df["Estq Seg"].fillna(0)
else:
    df["Estq Seg"] = 0

df = df.fillna(0)

# ========================= CONFIGURAÇÕES DE ALMOXARIFADO =========================
st.markdown("### ⚙️ Configuração dos Almoxarifados")

col_cfg1, col_cfg2, col_cfg3 = st.columns(3)

with col_cfg1:
    idx_demanda = opcoes_almox.index("Saldo Almox 3") if "Saldo Almox 3" in opcoes_almox else 0
    almox_demanda = st.selectbox(
        "Almox para **Demanda / Pedidos**",
        options=opcoes_almox,
        index=idx_demanda,
        help="Saldo deste almoxarifado será comparado com a demanda de pedidos"
    )

with col_cfg2:
    idx_seg = opcoes_almox.index("Saldo Almox 30") if "Saldo Almox 30" in opcoes_almox else min(1, len(opcoes_almox)-1)
    almox_estq_seg = st.selectbox(
        "Almox para **Estoque de Segurança**",
        options=opcoes_almox,
        index=idx_seg,
        help="Saldo deste almoxarifado será comparado com o estoque de segurança dos parâmetros"
    )

with col_cfg3:
    if tem_ordens:
        considerar_ordens = st.checkbox(
            "📋 Incluir Ordens em Aberto no Saldo Real",
            value=True,
            help="Soma as ordens de produção em aberto ao saldo disponível no cálculo do Saldo Real"
        )
    else:
        st.info("ℹ️ Ordens não carregadas")
        considerar_ordens = False

st.divider()

# ========================= CÁLCULOS =========================
df["Saldo vs Demanda"] = df[almox_demanda] - df["Demanda Pedido"]

saldo_real = df[almox_demanda].copy()
if considerar_ordens and tem_ordens:
    saldo_real = saldo_real + df["Qtde Ordens Abertas"]

df["Saldo Real"] = saldo_real - df["Demanda Pedido"] - df["Qtde Pendente OP"]
df["Abaixo Estq Seg"] = df[almox_estq_seg] < df["Estq Seg"]

# ========================= STATUS =========================
def status(row):
    if row["Saldo Real"] < 0:
        return "FALTA"
    if row["Abaixo Estq Seg"]:
        return "RISCO"
    if row["Demanda Pedido"] + row["Qtde Pendente OP"] >= row[almox_estq_seg] * 0.5:
        return "RISCO"
    return "OK"

df["Status"] = df.apply(status, axis=1)

# ========================= CARDS =========================
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

# ========================= LEGENDA =========================
col_leg1, col_leg2 = st.columns(2)
col_leg1.caption(f"📦 Demanda/Pedidos usando: `{almox_demanda}`")
col_leg2.caption(f"🔒 Estq. Segurança usando: `{almox_estq_seg}`" + (" + 📋 Ordens em Aberto" if considerar_ordens and tem_ordens else ""))

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

# ========================= COLUNAS EXIBIDAS =========================
colunas_exibir = ["Codigo", "Descricao", almox_demanda]
if almox_estq_seg != almox_demanda:
    colunas_exibir.append(almox_estq_seg)
colunas_exibir += ["Demanda Pedido", "Qtde Pendente OP"]
if tem_ordens:
    colunas_exibir.append("Qtde Ordens Abertas")
colunas_exibir += ["Saldo vs Demanda", "Saldo Real"]
if tem_parametros:
    colunas_exibir.append("Estq Seg")
colunas_exibir.append("Status")
colunas_exibir = [c for c in colunas_exibir if c in df_filtrado.columns]

# ========================= COR NA LINHA =========================
def cor(row):
    if row["Status"]=="FALTA":
        return ["background-color:#fecaca"]*len(row)
    if row["Status"]=="RISCO":
        return ["background-color:#fde68a"]*len(row)
    if row["Status"]=="OK":
        return ["background-color:#bbf7d0"]*len(row)
    return [""]*len(row)

st.dataframe(df_filtrado[colunas_exibir].style.apply(cor, axis=1), use_container_width=True)

# ========================= ALERTA ESTQ SEG =========================
if tem_parametros:
    df_abaixo = df[df["Abaixo Estq Seg"] & (df["Status"] != "FALTA")].copy()
    if not df_abaixo.empty:
        with st.expander(f"⚠️ {len(df_abaixo)} item(ns) abaixo do Estoque de Segurança ({almox_estq_seg})", expanded=True):
            cols_alerta = ["Codigo", "Descricao", almox_estq_seg, "Estq Seg", "Status"]
            cols_alerta = [c for c in cols_alerta if c in df_abaixo.columns]
            df_alerta = df_abaixo[cols_alerta].copy()
            df_alerta["Diferença"] = df_alerta[almox_estq_seg] - df_alerta["Estq Seg"]
            st.dataframe(df_alerta.style.apply(cor, axis=1), use_container_width=True)
    else:
        st.success("✅ Todos os itens estão acima do Estoque de Segurança.")
else:
    st.info("ℹ️ Importe o arquivo de Parâmetros para comparar com o Estoque de Segurança.")

# ========================= DOWNLOAD =========================
st.download_button("Baixar CSV", df_filtrado[colunas_exibir].to_csv(index=False), "pcp.csv")
