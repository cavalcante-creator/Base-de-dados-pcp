import streamlit as st
import pandas as pd
import os
from datetime import datetime
import pytz

st.set_page_config(page_title="Base de Dados PCP", layout="wide")

# ==========================================================

# FUNÇÃO DATA/HORA

# ==========================================================

def agora():
fuso = pytz.timezone("America/Sao_Paulo")
return datetime.now(fuso)

# ==========================================================

# MENU LATERAL

# ==========================================================

pagina = st.sidebar.radio(
"Escolha uma opção:",
[
"Upload Ordens de Fabricação",
"Upload Previsão Produção",
"Upload Parâmetros",
"Gerar Excel"
]
)

# ==========================================================

# UPLOAD ORDENS DE FABRICAÇÃO

# ==========================================================

if pagina == "Upload Ordens de Fabricação":

```
st.title("📦 Upload Ordens de Fabricação")

arquivo_of = st.file_uploader(
    "Selecione o arquivo de OF",
    type=["xlsx", "xls", "csv"]
)

if arquivo_of:
    st.success("Arquivo de OF carregado com sucesso!")
```

# ==========================================================

# UPLOAD PREVISÃO PRODUÇÃO

# ==========================================================

if pagina == "Upload Previsão Produção":

```
st.title("📈 Upload Previsão Produção")

arquivo_prev = st.file_uploader(
    "Selecione o arquivo de previsão",
    type=["xlsx", "xls", "csv"]
)

if arquivo_prev:
    st.success("Arquivo de previsão carregado com sucesso!")
```

# ==========================================================

# UPLOAD PARÂMETROS

# ==========================================================

if pagina == "Upload Parâmetros":

```
st.title("⚙️ Parâmetros Produção")

file = st.file_uploader(
    "Excel Parâmetros",
    type=["xls", "xlsx"],
    accept_multiple_files=False
)

if file:

    try:
        file.seek(0)
        df_raw = pd.read_excel(
            file,
            header=None,
            engine="openpyxl"
        )
        engine_excel = "openpyxl"

    except:
        file.seek(0)
        df_raw = pd.read_excel(
            file,
            header=None,
            engine="xlrd"
        )
        engine_excel = "xlrd"

    linha_header = None

    for i in range(len(df_raw)):
        linha = df_raw.iloc[i].astype(str).str.upper().str.strip().values

        if "COD ITEM" in linha:
            linha_header = i
            break

    if linha_header is None:
        st.error("Não foi possível localizar a linha de cabeçalho.")

    else:

        file.seek(0)

        df = pd.read_excel(
            file,
            header=linha_header,
            engine=engine_excel
        )

        df.columns = df.columns.astype(str).str.upper().str.strip()

        colunas_necessarias = [
            "COD ITEM",
            "DESC TECNICA",
            "UM",
            "LOTE MIN",
            "LOTE MAX",
            "LOTE MULT",
            "ESTQ SEG",
            "TEMP REP",
            "TEMP SEG",
            "AGRUP",
            "PLANEJADOR",
            "CONS MEDIO"
        ]

        colunas_encontradas = [
            c for c in colunas_necessarias if c in df.columns
        ]

        if len(colunas_encontradas) == 0:
            st.error("Nenhuma das colunas esperadas foi encontrada no arquivo.")

        else:

            df = df[colunas_encontradas]

            renomear = {
                "COD ITEM": "COD ITEM",
                "DESC TECNICA": "DESCRICAO",
                "UM": "UM",
                "LOTE MIN": "LOTE MIN",
                "LOTE MAX": "LOTE MAX",
                "LOTE MULT": "LOTE MULT",
                "ESTQ SEG": "ESTOQUE SEGURANCA",
                "TEMP REP": "TEMPO REPOSICAO",
                "TEMP SEG": "TEMPO SEGURANCA",
                "AGRUP": "AGRUPAMENTO",
                "PLANEJADOR": "PLANEJADOR",
                "CONS MEDIO": "CONSUMO MEDIO"
            }

            df = df.rename(columns=renomear)

            df["Data Processamento"] = agora().strftime("%d/%m/%Y")
            df["Hora Processamento"] = agora().strftime("%H:%M:%S")

            arquivo = "parametros.csv"

            if os.path.exists(arquivo):
                df_antigo = pd.read_csv(arquivo)
                df = pd.concat([df_antigo, df], ignore_index=True)

            df.to_csv(arquivo, index=False)

            st.success("Parâmetros carregados com sucesso!")
            st.dataframe(df, use_c
```
