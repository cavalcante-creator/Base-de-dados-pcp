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

# ========================= FUNÇÕES =========================
fuso = pytz.timezone("America/Sao_Paulo")

def agora():
    return datetime.now(fuso)

def exportar_excel_formatado(df, nome_aba="Dados"):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=nome_aba)
        ws = writer.book[nome_aba]

        ws.freeze_panes = "A2"

        header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
        header_font = Font(color="FFFFFF", bold=True)

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        if ws.max_row >= 2:
            ultima_coluna = get_column_letter(ws.max_column)
            tabela = Table(
                displayName="TabelaDados",
                ref=f"A1:{ultima_coluna}{ws.max_row}"
            )

            tabela.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium9",
                showRowStripes=True
            )

            ws.add_table(tabela)

    output.seek(0)
    return output.getvalue()

def botao_downloads(df):
    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            "📄 CSV",
            df.to_csv(index=False).encode("utf-8"),
            "dashboard_pcp.csv"
        )

    with col2:
        st.download_button(
            "📊 Excel",
            exportar_excel_formatado(df),
            "dashboard_pcp.xlsx"
        )

# ========================= FUNÇÃO GENÉRICA =========================
def achar_coluna(df, possiveis):
    for col in possiveis:
        if col in df.columns:
            return col
    return None

# ========================= DADOS =========================
try:
    saldo = pd.read_csv("saldo.csv")
    perfil = pd.read_csv("perfil.csv")
    previsao = pd.read_csv("previsao.csv")
    ordens = pd.read_csv("ordens.csv")
except:
    st.warning("Faça upload dos dados primeiro.")
    st.stop()

# ========================= IDENTIFICAR COLUNAS =========================
col_item_perfil = achar_coluna(perfil, ["Cd. Item","Item","Codigo","Cod Item"])
col_item_ordens = achar_coluna(ordens, ["Cd. Item","Item","Codigo","Cod Item"])

col_qtd_perfil = achar_coluna(perfil, ["Quantidade","Qtde","Qtd"])
col_qtd_ordens = achar_coluna(ordens, ["Quantidade","Qtde","Qtd"])

if not col_item_perfil or not col_qtd_perfil:
    st.error(f"Erro no perfil.csv → {list(perfil.columns)}")
    st.stop()

if not col_item_ordens or not col_qtd_ordens:
    st.error(f"Erro no ordens.csv → {list(ordens.columns)}")
    st.stop()

# ========================= TRATAR QUANTIDADE =========================
perfil[col_qtd_perfil] = (
    perfil[col_qtd_perfil].astype(str)
    .str.replace(".", "", regex=False)
    .str.replace(",", ".", regex=False)
    .astype(float)
)

ordens[col_qtd_ordens] = (
    ordens[col_qtd_ordens].astype(str)
    .str.replace(".", "", regex=False)
    .str.replace(",", ".", regex=False)
    .astype(float)
)

# ========================= BASES =========================
saldo = saldo.sort_values(by=["Data Processamento","Hora Processamento"], ascending=False).drop_duplicates("Codigo")
previsao = previsao.sort_values(by=["Data Processamento","Hora Processamento"], ascending=False).drop_duplicates("COD")

base = previsao[["COD","PRODUTO"]].copy()
base.columns = ["Codigo","Descricao"]

saldo_base = saldo[["Codigo","Saldo Total","Saldo Almox 3"]]

# ========================= DEMANDA =========================
dc = perfil[perfil["Tipo"]=="DC"] \
    .groupby(col_item_perfil)[col_qtd_perfil].sum().reset_index()

dc.columns = ["Codigo","Demanda Pedido"]

# ========================= ORDENS (OFA) =========================
op = ordens[ordens["Tipo"]=="OFA"] \
    .groupby(col_item_ordens)[col_qtd_ordens].sum().reset_index()

op.columns = ["Codigo","Qtde Pendente OP"]

# ========================= MERGE =========================
df = base.merge(saldo_base,on="Codigo",how="left")
df = df.merge(dc,on="Codigo",how="left")
df = df.merge(op,on="Codigo",how="left")

df = df.fillna(0)

# ========================= LÓGICA ORIGINAL =========================
df["Saldo vs Demanda"] = df["Saldo Almox 3"] - df["Demanda Pedido"]

def status(row):
    if row["Saldo vs Demanda"] < 0:
        return "FALTA"
    if row["Demanda Pedido"] >= row["Saldo Almox 3"] * 0.5:
        return "RISCO"
    return "OK"

df["Status"] = df.apply(status, axis=1)

# ========================= NOVO (SEM INTERFERIR) =========================
df["Saldo Real"] = (
    df["Saldo Almox 3"]
    - df["Demanda Pedido"]
    - df["Qtde Pendente OP"]
)

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

# ========================= DOWNLOAD =========================
botao_downloads(df_filtrado)
