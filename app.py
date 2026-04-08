import streamlit as st
import pandas as pd
from datetime import datetime

# 🔧 CONFIG
st.set_page_config(page_title="Dashboard PCP", layout="wide")

# 🎨 ESTILO
st.markdown("""
<style>
.block-container { padding-top: 1rem; }
.header { font-size: 28px; font-weight: bold; }
.subheader { font-size: 14px; color: gray; }
</style>
""", unsafe_allow_html=True)

# 🏷️ TÍTULO
st.markdown('<div class="header">📊 Dashboard PCP</div>', unsafe_allow_html=True)

# 🕒 DATA/HORA
agora = datetime.now()
st.markdown(
    f'<div class="subheader">Atualizado em: {agora.strftime("%d/%m/%Y %H:%M:%S")}</div>',
    unsafe_allow_html=True
)

st.divider()

# 📂 FUNÇÃO DE LEITURA
def ler_arquivo(file):
    if file.name.endswith(".csv"):
        return pd.read_csv(file)
    elif file.name.endswith(".xls"):
        return pd.read_excel(file, engine="xlrd")
    else:
        return pd.read_excel(file, engine="openpyxl")

# 📁 UPLOADS
col1, col2, col3 = st.columns(3)

with col1:
    file_saldo = st.file_uploader("📦 Upload Saldo", type=["xlsx", "xls", "csv"])

with col2:
    file_ordens = st.file_uploader("🏭 Upload Ordens", type=["xlsx", "xls", "csv"])

with col3:
    file_perfil = st.file_uploader("📋 Upload Perfil", type=["xlsx", "xls", "csv"])

# 🔄 PROCESSAMENTO
if file_saldo and file_ordens and file_perfil:

    try:
        df_saldo = ler_arquivo(file_saldo)
        df_ordens = ler_arquivo(file_ordens)
        df_perfil = ler_arquivo(file_perfil)

        st.success("✅ Arquivos carregados com sucesso!")

        # 🔧 PADRONIZAÇÃO (ajusta nomes conforme seu layout)
        df_saldo.columns = df_saldo.columns.str.strip()
        df_ordens.columns = df_ordens.columns.str.strip()
        df_perfil.columns = df_perfil.columns.str.strip()

        # ⚠️ GARANTE COLUNAS
        # (ajuste aqui se seus nomes forem diferentes)
        df_saldo = df_saldo.rename(columns={
            "Codigo": "Codigo",
            "Saldo Total": "Saldo"
        })

        df_ordens = df_ordens.rename(columns={
            "Cód. Item": "Codigo",
            "Qtde.": "Ordem_Qtd",
            "Dt. Fim": "Data"
        })

        df_perfil = df_perfil.rename(columns={
            "Item": "Codigo",
            "Quantidade": "Demanda",
            "Tipo": "Tipo"
        })

        # 📅 CONVERTE DATA
        if "Data" in df_ordens.columns:
            df_ordens["Data"] = pd.to_datetime(df_ordens["Data"], errors="coerce")

        # 🔗 MERGE
        df = df_ordens.merge(df_saldo, on="Codigo", how="left")
        df = df.merge(df_perfil, on="Codigo", how="left")

        # 📊 MÉTRICAS
        total_ordens = df["Ordem_Qtd"].sum()
        total_saldo = df["Saldo"].sum()
        total_demanda = df["Demanda"].sum()

        col1, col2, col3 = st.columns(3)

        col1.metric("📦 Produção Total", f"{total_ordens:,.0f}")
        col2.metric("🏪 Saldo Total", f"{total_saldo:,.0f}")
        col3.metric("📉 Demanda Total", f"{total_demanda:,.0f}")

        st.divider()

        # 🎛️ FILTROS
        col1, col2 = st.columns(2)

        with col1:
            itens = st.multiselect("Filtrar por Item", df["Codigo"].dropna().unique())

        with col2:
            tipos = st.multiselect("Filtrar por Tipo", df["Tipo"].dropna().unique())

        # 🔍 APLICA FILTROS
        df_filtrado = df.copy()

        if itens:
            df_filtrado = df_filtrado[df_filtrado["Codigo"].isin(itens)]

        if tipos:
            df_filtrado = df_filtrado[df_filtrado["Tipo"].isin(tipos)]

        # 📊 RESULTADO
        st.subheader("📋 Base Consolidada")
        st.dataframe(df_filtrado, use_container_width=True)

        # 📥 DOWNLOAD
        def convert(df):
            return df.to_csv(index=False).encode("utf-8")

        st.download_button(
            "📥 Baixar Base Consolidada",
            convert(df_filtrado),
            "pcp_consolidado.csv",
            "text/csv"
        )

    except Exception as e:
        st.error("❌ Erro no processamento")
        st.code(str(e))

else:
    st.warning("⚠️ Envie os 3 arquivos para visualizar o dashboard.")

# 🔚 RODAPÉ
st.divider()
st.caption("PCP Dashboard • Versão Completa 🚀")
