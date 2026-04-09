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
# CSV ROBUSTO
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

# ==========================================================
# SALVAR CSV
# ==========================================================
def salvar_csv(df, nome):
    if os.path.exists(nome):
        antigo = ler_csv_seguro(nome)
        df = pd.concat([antigo, df], ignore_index=True)
    df.to_csv(nome, index=False)

# ==========================================================
# LIMPAR BASE
# ==========================================================
def limpar_base():
    for arq in ["saldo.csv","perfil.csv","ordens.csv","previsao.csv","parametros.csv"]:
        if os.path.exists(arq):
            os.remove(arq)

# ==========================================================
# ABAS (TOPO)
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
    st.title("📦 Saldo")

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
                    dados[cod]={"Codigo":cod,"Saldo Total":0,"Saldo Almox 3":0}
                    continue

                if "ALMOXARIFADO" in linha.upper() and cod:
                    nums=re.findall(r'[\d\.]+\,\d+',linha)
                    if nums:
                        v=float(nums[-1].replace(".","").replace(",","."))
                        dados[cod]["Saldo Total"]+=v
                        dados[cod]["Saldo Almox 3"]=v

            df=pd.DataFrame(dados.values())
            df["Data Processamento"]=agora().strftime("%d/%m/%Y")

            salvar_csv(df,"saldo.csv")
            st.dataframe(df)

# ==========================================================
# PERFIL
# ==========================================================
with abas[1]:
    st.title("📊 Perfil")

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
# ORDENS
# ==========================================================
with abas[2]:
    st.title("📄 Ordens")

    file = st.file_uploader("CSV", type=["csv"])

    if file:
        conteudo=file.read().decode("latin-1",errors="ignore")
        df=ler_csv_seguro(StringIO(conteudo))
        df["Data Processamento"]=agora().strftime("%d/%m/%Y")

        salvar_csv(df,"ordens.csv")
        st.dataframe(df)

# ==========================================================
# PREVISÃO
# ==========================================================
with abas[3]:
    st.title("📅 Previsão")

    file = st.file_uploader("Excel", type=["xlsx"])

    if file:
        df=pd.read_excel(file)
        df.columns=df.columns.astype(str).str.upper()

        col_cod=[c for c in df.columns if "COD" in c][0]
        col_desc=[c for c in df.columns if "PROD" in c or "DESC" in c][0]

        df=df[[col_cod,col_desc]]
        df.columns=["COD","DESCRICAO"]

        df["Data Processamento"]=agora().strftime("%d/%m/%Y")

        salvar_csv(df,"previsao.csv")
        st.dataframe(df)

# ==========================================================
# PARÂMETROS (AJUSTADO PARA SUA PLANILHA)
# ==========================================================
with abas[4]:
    st.title("⚙️ Parâmetros")

    file = st.file_uploader("Arquivo Parâmetros", type=["xlsx","csv"])

    if file:
        try:
            df=pd.read_excel(file)

            df.columns=df.columns.astype(str).str.upper()

            col_cod = [c for c in df.columns if "COD" in c][0]
            col_seg = [c for c in df.columns if "ESTQ" in c][0]

            df = df[[col_cod, col_seg]]
            df.columns = ["COD","ESTQ SEG"]

            df["ESTQ SEG"]=pd.to_numeric(df["ESTQ SEG"],errors="coerce").fillna(0)

            df["Data Processamento"]=agora().strftime("%d/%m/%Y")

            salvar_csv(df,"parametros.csv")

            st.success("Parâmetros carregados!")
            st.dataframe(df)

        except Exception as e:
            st.error(f"Erro: {e}")

# ==========================================================
# BASE
# ==========================================================
with abas[5]:
    st.title("📋 Base")

    if st.button("🗑 Limpar Base"):
        limpar_base()

    for arq in ["saldo.csv","perfil.csv","ordens.csv","previsao.csv","parametros.csv"]:
        st.subheader(arq)
        if os.path.exists(arq):
            st.dataframe(ler_csv_seguro(arq))

# ==========================================================
# DASHBOARD FINAL
# ==========================================================
with abas[6]:
    st.title("📊 Dashboard PCP")

    saldo=ler_csv_seguro("saldo.csv")
    perfil=ler_csv_seguro("perfil.csv")
    previsao=ler_csv_seguro("previsao.csv")

    datas=saldo["Data Processamento"].unique()
    data_sel=st.selectbox("📅 Data",datas)

    saldo=saldo[saldo["Data Processamento"]==data_sel]
    perfil=perfil[perfil["Data Processamento"]==data_sel]
    previsao=previsao[previsao["Data Processamento"]==data_sel]

    base=previsao.rename(columns={"COD":"Codigo","DESCRICAO":"Descricao"})

    perfil["Quantidade"]=perfil["Quantidade"].astype(str)\
        .str.replace(".","").str.replace(",",".").astype(float)

    # DEMANDA DC
    dc=perfil[perfil["Tipo"]=="DC"].groupby("Item")["Quantidade"].sum().reset_index()
    dc.columns=["Codigo","Demanda Pedido"]

    # DEMANDA DP TOTAL
    dp=perfil[perfil["Tipo"]=="DP"].groupby("Item")["Quantidade"].sum().reset_index()
    dp.columns=["Codigo","Demanda DP"]

    # DEMANDA DP SEMANA
    semana=str(datetime.now().isocalendar()[1]).zfill(2)
    perfil["Semana"]=pd.to_datetime(perfil["Data Processamento"],dayfirst=True).dt.isocalendar().week.astype(str).str.zfill(2)

    dp_sem=perfil[(perfil["Tipo"]=="DP")&(perfil["Semana"]==semana)]\
        .groupby("Item")["Quantidade"].sum().reset_index()

    dp_sem.columns=["Codigo","DP Semana"]

    df=base.merge(saldo,on="Codigo",how="left")\
           .merge(dc,on="Codigo",how="left")\
           .merge(dp,on="Codigo",how="left")\
           .merge(dp_sem,on="Codigo",how="left")

    df=df.fillna(0)

    df["Saldo vs Demanda"]=df["Saldo Almox 3"]-df["Demanda Pedido"]

    def status(x):
        if x<0: return "🔴 FALTA"
        elif x<50: return "🟡 RISCO"
        else: return "🟢 OK"

    df["Status"]=df["Saldo vs Demanda"].apply(status)

    st.dataframe(df, use_container_width=True)
