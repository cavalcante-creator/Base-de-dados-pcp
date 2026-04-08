import streamlit as st
import pandas as pd
from datetime import datetime

# 🔧 CONFIG DA PÁGINA
st.set_page_config(
    page_title="Base de Dados PCP",
    layout="wide"
)

# 🧠 HEADER
st.markdown("""
<style>
.block-container { padding-top: 1rem; }
.header {
    font-size: 28px;
    font-weight: bold;
}
.subheader {
    font-size: 16px;
    color: gray;
}
</style>
""", unsafe_allow_html=True)

# 🏷️ TÍTULO
st.markdown('<div class="header">📊 Base de Dados PCP</div>', unsafe_allow_html=True)

# 🕒 DATA E HORA
agora = datetime.now()
st.markdown(
    f'<div class="subheader">Atualizado em: {agora.strftime("%d/%m/%Y %H:%M:%S")}</div>',
    unsafe_allow_html=True
)

st.divider()

# 📂 UPLOAD
file = st.file_uploader(
    "📁 Upload saldo produção",
    type=["xlsx", "xls", "csv"]
)

# 🔄 PROCESSAMENTO
if file is not None:
    try:
        # 📌 IDENTIFICA TIPO
        if file.name.endswith(".csv"):
            df_raw = pd.read_csv(file, header=None)

        elif file.name.endswith(".xls"):
            df_raw = pd.read_excel(file, header=None, engine="xlrd")

        else:  # .xlsx
            df_raw = pd.read_excel(file, header=None, engine="openpyxl")

        # ✅ SUCESSO
        st.success("Arquivo carregado com sucesso!")

        # 👀 VISUALIZAÇÃO
        st.subheader("🔍 Pré-visualização dos dados")
        st.dataframe(df_raw, use_container_width=True)

        # 📊 INFO RÁPIDA
        col1, col2, col3 = st.columns(3)

        with col1:
            st.metric("Linhas", df_raw.shape[0])

        with col2:
            st.metric("Colunas", df_raw.shape[1])

        with col3:
            st.metric("Arquivo", file.name)

        # 📥 DOWNLOAD (excel tratado)
        def convert_excel(df):
            return df.to_csv(index=False).encode("utf-8")

        st.download_button(
            "📥 Baixar dados tratados (CSV)",
            convert_excel(df_raw),
            file_name="dados_tratados.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error("❌ Erro ao ler o arquivo")
        st.code(str(e))

else:
    st.warning("⚠️ Por favor, envie um arquivo para começar.")

# 🔚 RODAPÉ
st.divider()
st.caption("PCP Dashboard • Desenvolvido em Streamlit")
