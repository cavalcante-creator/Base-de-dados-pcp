import streamlit as st
import pdfplumber
import pandas as pd
import re
from datetime import datetime
import os
from io import StringIO
import pytz

st.set_page_config(page_title="PCP Produção", layout="wide")

fuso = pytz.timezone("America/Sao_Paulo")

def agora():
    return datetime.now(fuso)

# ==========================================================
# FUNÇÕES BASE
# ==========================================================
def ler_csv_seguro(file):
    try:
        return pd.read_csv(file, sep=None, engine="python", encoding="latin-1", on_bad_lines="skip")
    except:
        try:
            return pd.read_csv(file, sep=";", encoding="latin-1", on_bad_lines="skip")
        except:
            try:
                return pd.read_csv(file, sep=",", encoding="latin-1", on_bad_lines="skip")
            except:
                return pd.DataFrame()

def salvar_csv(df, nome):
    if os.path.exists(nome):
        antigo = ler_csv_seguro(nome)
        df = pd.concat([antigo, df], ignore_index=True)
    df.to_csv(nome, index=False)

def limpar_base():
    for arq in ["saldo.csv","perfil.csv","ordens.csv","previsao.csv","parametros.csv"]:
        if os.path.exists(arq):
            os.remove(arq)

# ==========================================================
# ABAS
# ==========================================================
abas = st.tabs([
    "📦 Saldo",
    "📊 Perfil",
    "📄 Ordens",
    "📅 Previsão",
    "⚙️ Parâmetros",
    "📋 Base",
    "📊 Dashboard"
])

# ==========================================================
# SALDO
# ==========================================================
with abas[0]:
    file = st.file_uploader("PDF Saldo", type=["pdf"])

    if file:
        with open("saldo.pdf","wb") as f:
            f.write(file.read())

        if st.button("Processar Saldo"):
            linhas=[]
            with pdfplumber.open("saldo.pdf") as pdf:
                for p in pdf.pages:
                    texto=p.extract_text()
                    if texto:
                        linhas.extend(texto.split("\n"))

            dados={}
            cod=None

            for linha in linhas:
                m=re.search(r'\b([A-Z]{1,3}\d{3,5})\b',linha)
                if m:
                    cod=m.group(1)
                    if cod not in dados:
                        dados[cod]={"Codigo":cod,"Saldo Total":0,"Saldo Almox 3":0}
                    continue

                if "ALMOXARIFADO" in linha.upper() and cod:
                    nums=re.findall(r'[\d\.]+\,\d+',linha)
                    if nums:
                        v=float(nums[-1].replace(".","").replace(",","."))
                        dados[cod]["Saldo Total"]+=v
                        dados[cod]["Saldo Almox 3"]+=v

            df=pd.DataFrame(dados.values())
            df["Data Processamento"]=agora().strftime("%d/%m/%Y")

            salvar_csv(df,"saldo.csv")
            st.dataframe(df)

# ==========================================================
# PERFIL
# ==========================================================
with abas[1]:
    file = st.file_uploader("PDF Perfil", type=["pdf"])

    if file:
        with open("perfil.pdf","wb") as f:
            f.write(file.read())

        if st.button("Processar Perfil"):
            dados=[]
            item=""

            regex=re.compile(r'(DD|DC|DP).*?(-?[\d,.]+)')

            with pdfplumber.open("perfil.pdf") as pdf:
                for p in pdf.pages:
                    texto=p.extract_text()
                    if texto:
                        for l in texto.split("\n"):
                            m_item=re.search(r'Item:\s*(\S+)',l)
                            if m_item:
                                item=m_item.group(1)

                            m=regex.search(l)
                            if m:
                                dados.append({
                                    "Item":item,
                                    "Tipo":m.group(1),
                                    "Quantidade":m.group(2)
                                })

            df=pd.DataFrame(dados)
            df["Data Processamento"]=agora().strftime("%d/%m/%Y")

            salvar_csv(df,"perfil.csv")
            st.dataframe(df)

# ==========================================================
# PREVISÃO
# ==========================================================
with abas[3]:
    file = st.file_uploader("Excel Previsão", type=["xlsx","xls"])

    if file:
        df=pd.read_excel(file)
        df.columns=df.columns.astype(str).str.upper()

        col_cod = next((c for c in df.columns if "COD" in c), None)
        col_desc = next((c for c in df.columns if "DESC" in c or "PROD" in c), None)

        if col_cod is None or col_desc is None:
            st.error("Colunas não encontradas")
            st.stop()

        df=df[[col_cod,col_desc]]
        df.columns=["Codigo","Descricao"]

        df=df.drop_duplicates(subset=["Codigo"])

        df["Data Processamento"]=agora().strftime("%d/%m/%Y")

        salvar_csv(df,"previsao.csv")
        st.dataframe(df)

# ==========================================================
# PARÂMETROS
# ==========================================================
with abas[4]:
    file = st.file_uploader("Parâmetros", type=None)

    if file:
        try:
            try:
                df = pd.read_excel(file)
            except:
                df = ler_csv_seguro(file)

            df.columns=df.columns.astype(str).str.upper()

            col_cod = next((c for c in df.columns if "COD" in c), None)
            col_seg = next((c for c in df.columns if "ESTQ" in c or "SEG" in c), None)

            if col_cod is None or col_seg is None:
                st.error("Colunas não encontradas")
                st.write(df.columns.tolist())
                st.stop()

            df=df[[col_cod,col_seg]]
            df.columns=["Codigo","Estq Seg"]

            df["Estq Seg"]=pd.to_numeric(df["Estq Seg"],errors="coerce").fillna(0)

            df["Data Processamento"]=agora().strftime("%d/%m/%Y")

            salvar_csv(df,"parametros.csv")
            st.dataframe(df)

        except Exception as e:
            st.error(e)

# ==========================================================
# DASHBOARD
# ==========================================================
with abas[6]:
    st.title("📊 Dashboard PCP")

    if not os.path.exists("saldo.csv"):
        st.warning("Faça upload dos dados")
        st.stop()

    saldo=ler_csv_seguro("saldo.csv")
    perfil=ler_csv_seguro("perfil.csv")
    previsao=ler_csv_seguro("previsao.csv")

    data_sel=saldo["Data Processamento"].max()

    saldo=saldo[saldo["Data Processamento"]==data_sel]
    perfil=perfil[perfil["Data Processamento"]==data_sel]
    previsao=previsao[previsao["Data Processamento"]==data_sel]

    perfil["Quantidade"]=perfil["Quantidade"].astype(str)\
        .str.replace(".","").str.replace(",",".").astype(float)

    dc=perfil[perfil["Tipo"]=="DC"].groupby("Item")["Quantidade"].sum().reset_index()
    dc.columns=["Codigo","Demanda Pedido"]

    dp=perfil[perfil["Tipo"]=="DP"].groupby("Item")["Quantidade"].sum().reset_index()
    dp.columns=["Codigo","Demanda DP"]

    semana=str(datetime.now().isocalendar()[1]).zfill(2)

    perfil["Semana"]=pd.to_datetime(perfil["Data Processamento"],dayfirst=True)\
        .dt.isocalendar().week.astype(str).str.zfill(2)

    dp_sem=perfil[(perfil["Tipo"]=="DP")&(perfil["Semana"]==semana)]\
        .groupby("Item")["Quantidade"].sum().reset_index()

    dp_sem.columns=["Codigo","Demanda DP Semana"]

    df=previsao.merge(saldo,on="Codigo",how="left")\
        .merge(dc,on="Codigo",how="left")\
        .merge(dp,on="Codigo",how="left")\
        .merge(dp_sem,on="Codigo",how="left")

    df=df.fillna(0)

    # REMOVE DUPLICADOS 🔥
    df=df.groupby(["Codigo","Descricao"],as_index=False).sum()

    df["Saldo vs Demanda"]=df["Saldo Almox 3"]-df["Demanda Pedido"]

    def status(x):
        if x<0: return "FALTA"
        elif x<50: return "RISCO"
        else: return "OK"

    df["Status"]=df["Saldo vs Demanda"].apply(status)

    # CARDS 🔥
    col1,col2,col3=st.columns(3)

    col1.metric("🔴 FALTA", len(df[df["Status"]=="FALTA"]))
    col2.metric("🟡 RISCO", len(df[df["Status"]=="RISCO"]))
    col3.metric("🟢 OK", len(df[df["Status"]=="OK"]))

    # TABELA FINAL
    df_final=df[[
        "Codigo",
        "Descricao",
        "Saldo Total",
        "Saldo Almox 3",
        "Demanda Pedido",
        "Status",
        "Demanda DP",
        "Demanda DP Semana"
    ]]

    st.dataframe(df_final, use_container_width=True)
