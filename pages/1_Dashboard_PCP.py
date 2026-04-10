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

# =========================
# ESTILO
# =========================
st.markdown("""
<style>
.main {background: linear-gradient(180deg, #f7fafc 0%, #eef4f7 100%);}
.block-container {padding-top: 2rem; padding-bottom: 2rem;}
h1, h2, h3 {color: #12344d; font-weight: 700;}

div[data-testid="stMetric"] {
    background: white;
    border: 1px solid #dbe7f0;
    border-radius: 18px;
    padding: 14px;
    box-shadow: 0 4px 14px rgba(18, 52, 77, 0.08);
}

.custom-card {
    background: white;
    padding: 20px;
    border-radius: 18px;
    border: 1px solid #dbe7f0;
    box-shadow: 0 6px 20px rgba(18, 52, 77, 0.08);
    margin-bottom: 16px;
}
.small-title {font-size: 14px; color: #5b7488;}
.big-number {font-size: 32px; font-weight: 700; color: #12344d;}
</style>
""", unsafe_allow_html=True)

# =========================
# DATA
# =========================
fuso = pytz.timezone("America/Sao_Paulo")
def agora():
    return datetime.now(fuso)

# =========================
# EXPORTAÇÃO
# =========================
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
            cell.alignment = Alignment(horizontal="center")

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
            "dashboard_pcp.csv"
        )

    with col2:
        st.download_button(
            "Baixar Excel",
            exportar_excel_formatado(df),
            "dashboard_pcp.xlsx"
        )

# =========================
# CABEÇALHO
# =========================
st.markdown("""
<div class="custom-card">
<div class="small-title">Painel Gerencial</div>
<div class="big-number">Dashboard PCP</div>
<div style="color:#5b7488;">Visão rápida dos itens</div>
</div>
""", unsafe_allow_html=True)

# =========================
# LEITURA
# =========================
try:
    saldo = pd.read_csv("saldo.csv")
    perfil = pd.read_csv("perfil.csv")
    previsao = pd.read_csv("previsao.csv")
except:
    st.warning("Faça upload dos dados primeiro.")
    st.stop()

saldo = saldo.sort_values(by=["Data Processamento","Hora Processamento"], ascending=False).drop_duplicates("Codigo")
previsao = previsao.drop_duplicates("COD")

base = previsao[["COD","PRODUTO"]].copy()
base.columns = ["Codigo","Descricao"]

saldo_base = saldo[["Codigo","Saldo Total","Saldo Almox 3"]]

perfil["Quantidade"] = (
    perfil["Quantidade"].astype(str)
    .str.replace(".","",regex=False)
    .str.replace(",",".",regex=False)
    .astype(float)
)

dc = perfil[perfil["Tipo"]=="DC"].groupby("Item")["Quantidade"].sum().reset_index()
dc.columns = ["Codigo","Demanda Pedido"]

# =========================
# MERGE
# =========================
df = base.merge(saldo_base,on="Codigo",how="left")
df = df.merge(dc,on="Codigo",how="left")
df = df.fillna(0)

df["Saldo vs Demanda"] = df["Saldo Almox 3"] - df["Demanda Pedido"]

# =========================
# STATUS AUTO
# =========================
def status(row):
    if row["Saldo vs Demanda"] < 0:
        return "FALTA"
    if row["Demanda Pedido"] >= row["Saldo Almox 3"] * 0.5:
        return "RISCO"
    return "OK"

df["Status Auto"] = df.apply(status, axis=1)

# =========================
# STATUS MANUAL
# =========================
if "Status Manual" not in st.session_state:
    st.session_state["Status Manual"] = {}

df["Status Manual"] = df["Codigo"].map(st.session_state["Status Manual"])
df["Status Final"] = df["Status Manual"].fillna(df["Status Auto"])

# =========================
# CARDS
# =========================
c1,c2,c3,c4 = st.columns(4)
c1.metric("Total",len(df))
c2.metric("Falta",(df["Status Final"]=="FALTA").sum())
c3.metric("Risco",(df["Status Final"]=="RISCO").sum())
c4.metric("OK",(df["Status Final"]=="OK").sum())

st.markdown("---")

# =========================
# FILTROS
# =========================
status_sel = st.multiselect("Status",["FALTA","RISCO","OK"],default=["FALTA","RISCO","OK"])
busca = st.text_input("Buscar")

df_f = df[df["Status Final"].isin(status_sel)]

if busca:
    df_f = df_f[
        df_f["Codigo"].astype(str).str.contains(busca,case=False) |
        df_f["Descricao"].str.contains(busca,case=False)
    ]

# =========================
# GRÁFICO
# =========================
st.bar_chart(df["Status Final"].value_counts())

# =========================
# TABELA EDITÁVEL
# =========================
df_edit = st.data_editor(
    df_f,
    column_config={
        "Status Manual": st.column_config.SelectboxColumn(
            options=["FALTA","RISCO","OK"]
        )
    },
    use_container_width=True
)

# SALVAR
for _,row in df_edit.iterrows():
    if pd.notna(row["Status Manual"]):
        st.session_state["Status Manual"][row["Codigo"]] = row["Status Manual"]

# =========================
# DOWNLOAD
# =========================
botao_downloads(df_edit)
