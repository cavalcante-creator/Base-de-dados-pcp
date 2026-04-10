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
        width: 100%;
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

    .card {
        background: white;
        padding: 18px;
        border-radius: 18px;
        border: 1px solid #dbe7f0;
        box-shadow: 0 6px 20px rgba(18, 52, 77, 0.08);
        text-align: center;
        margin-bottom: 12px;
    }

    .card-title {
        font-size: 14px;
        color: #5b7488;
        margin-bottom: 8px;
    }

    .card-value {
        font-size: 28px;
        font-weight: bold;
        color: #12344d;
    }
</style>
""", unsafe_allow_html=True)

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

        if ws.max_row >= 2 and ws.max_column >= 1:
            ultima_coluna = get_column_letter(ws.max_column)
            nome_tabela = re.sub(r"\W+", "", nome_aba) or "Dados"

            tabela = Table(
                displayName=f"Tabela{nome_tabela[:20]}",
                ref=f"A1:{ultima_coluna}{ws.max_row}"
            )

            tabela.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium9",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False
            )

            ws.add_table(tabela)

    output.seek(0)
    return output.getvalue()


def botao_downloads(df, nome_base, nome_aba):
    col_csv, col_excel = st.columns(2)

    with col_csv:
        st.download_button(
            label=f"Baixar {nome_base} CSV",
            data=df.to_csv(index=False).encode("utf-8"),
            file_name=f"{nome_base}.csv",
            mime="text/csv"
        )

    with col_excel:
        st.download_button(
            label=f"Baixar {nome_base} Excel",
            data=exportar_excel_formatado(df, nome_aba),
            file_name=f"{nome_base}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


st.title("Dashboard PCP")

st.info("O dashboard sempre utiliza automaticamente o último processamento disponível na base de dados.")

previsao = st.session_state.get("previsao_df", pd.DataFrame())
saldo = st.session_state.get("saldo_df", pd.DataFrame())
perfil = st.session_state.get("perfil_df", pd.DataFrame())

if previsao.empty or saldo.empty or perfil.empty:
    st.warning("Faça o processamento dos arquivos antes de acessar o dashboard.")
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

botao_downloads(df_filtrado, "dashboard_pcp", "Dashboard_PCP")
