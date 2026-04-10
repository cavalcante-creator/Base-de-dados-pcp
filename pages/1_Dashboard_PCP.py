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
.main {
    background: linear-gradient(180deg, #f7fafc 0%, #eef4f7 100%);
}

.card {
    border-radius: 16px;
    padding: 18px;
    text-align: center;
    color: white;
    font-weight: 600;
    transition: 0.2s;
}

.card:hover {
    transform: scale(1.04);
}

.total { background: #1d4ed8; }
.falta { background: #dc2626; }
.risco { background: #f59e0b; }
.ok { background: #16a34a; }

.selected {
    border: 3px solid black;
}
</style>
""", unsafe_allow_html=True)

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
            "Baixar CSV",
            df.to_csv(index=False).encode("utf-8"),
            "dashboard.csv"
        )

    with col2:
        st.download_button(
            "Baixar Excel",
            exportar_excel_formatado(df),
            "dashboard.xlsx"
        )

# ========================= HEADER =========================
st.markdown("<h2>Dashboard PCP</h2>", unsafe_allow_html=True)

# ========================= DADOS =========================
try:
    saldo = pd.read_csv("saldo.csv")
    perfil = pd.read_csv("perfil.csv")
    previsao = pd.read_csv("previsao.csv")
except:
    st.warning("Faça upload dos dados primeiro.")
    st.stop()

saldo = saldo.sort_values(by=["Data Processamento", "Hora Processamento"], ascending=False).drop_duplicates("Codigo")
previsao = previsao.sort_values(by=["Data Processamento", "Hora Processamento"], ascending=False).drop_duplicates("COD")

base = previsao[["COD", "PRODUTO"]].copy()
base.columns = ["Codigo", "Descricao"]

saldo_base = saldo[["Codigo", "Saldo Total", "Saldo Almox 3"]]

perfil["Quantidade"] = (
    perfil["Quantidade"].astype(str)
    .str.replace(".", "", regex=False)
    .str.replace(",", ".", regex=False)
    .astype(float)
)

# ========================= DEMANDAS =========================
dc = perfil[perfil["Tipo"] == "DC"].groupby("Item")["Quantidade"].sum().reset_index()
dc.columns = ["Codigo", "Demanda Pedido"]

op = perfil[perfil["Tipo"] == "OP"].groupby("Item")["Quantidade"].sum().reset_index()
op.columns = ["Codigo", "Qtde Pendente OP"]

# ========================= MERGE =========================
df = base.merge(saldo_base, on="Codigo", how="left")
df = df.merge(dc, on="Codigo", how="left")
df = df.merge(op, on="Codigo", how="left")

df = df.fillna(0)

df["Saldo vs Demanda"] = df["Saldo Almox 3"] - df["Demanda Pedido"]

df["Saldo Real"] = (
    df["Saldo Almox 3"]
    - df["Demanda Pedido"]
    - df["Qtde Pendente OP"]
)

# ========================= STATUS =========================
def status(row):
    if row["Saldo Real"] < 0:
        return "FALTA"
    if row["Demanda Pedido"] + row["Qtde Pendente OP"] >= row["Saldo Almox 3"] * 0.5:
        return "RISCO"
    return "OK"

df["Status"] = df.apply(status, axis=1)

# ========================= FILTRO =========================
if "filtro_status" not in st.session_state:
    st.session_state.filtro_status = "TODOS"

col1, col2, col3, col4 = st.columns(4)

def card(col, titulo, valor, tipo):
    selecionado = st.session_state.filtro_status == tipo
    classe = f"card {tipo.lower()} {'selected' if selecionado else ''}"
    if col.button(f"{titulo}\n{valor}"):
        st.session_state.filtro_status = tipo

    col.markdown(f"<div class='{classe}'>{titulo}<br><h2>{valor}</h2></div>", unsafe_allow_html=True)

card(col1, "TOTAL", len(df), "TODOS")
card(col2, "FALTA", int((df["Status"]=="FALTA").sum()), "FALTA")
card(col3, "RISCO", int((df["Status"]=="RISCO").sum()), "RISCO")
card(col4, "OK", int((df["Status"]=="OK").sum()), "OK")

# ========================= FILTRAGEM =========================
if st.session_state.filtro_status == "TODOS":
    df_filtrado = df.copy()
else:
    df_filtrado = df[df["Status"] == st.session_state.filtro_status]

# ========================= BUSCA =========================
busca = st.text_input("Buscar")

if busca:
    df_filtrado = df_filtrado[
        df_filtrado["Codigo"].astype(str).str.contains(busca, case=False) |
        df_filtrado["Descricao"].astype(str).str.contains(busca, case=False)
    ]

# ========================= COR NA TABELA =========================
def cor_linha(row):
    if row["Status"] == "FALTA":
        return ["background-color: #fee2e2"] * len(row)
    elif row["Status"] == "RISCO":
        return ["background-color: #fef3c7"] * len(row)
    elif row["Status"] == "OK":
        return ["background-color: #dcfce7"] * len(row)
    return [""] * len(row)

df_styled = df_filtrado.style.apply(cor_linha, axis=1)

# ========================= TABELA =========================
st.dataframe(df_styled, use_container_width=True)

# ========================= DOWNLOAD =========================
botao_downloads(df_filtrado)
