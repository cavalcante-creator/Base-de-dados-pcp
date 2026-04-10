import streamlit as st
import pandas as pd

st.set_page_config(page_title="Dashboard PCP", layout="wide")

st.title("Dashboard PCP")

arquivo = st.file_uploader("Selecione a base Excel", type=["xlsx"])

if arquivo is not None:
    df = pd.read_excel(arquivo)

    # Garantir que colunas existam
    colunas_necessarias = [
        "Codigo",
        "Descricao",
        "Saldo Almox 1",
        "Saldo Almox 2",
        "Saldo Almox 3",
        "Saldo Almox 4",
        "Saldo Total",
        "Demanda",
        "Produzir"
    ]

    for col in colunas_necessarias:
        if col not in df.columns:
            df[col] = 0

    # Converter colunas numéricas
    colunas_numericas = [
        "Saldo Almox 1",
        "Saldo Almox 2",
        "Saldo Almox 3",
        "Saldo Almox 4",
        "Saldo Total",
        "Demanda",
        "Produzir"
    ]

    for col in colunas_numericas:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # Pegar sempre o último processamento
    if "DataHora_Processamento" in df.columns:
        df["DataHora_Processamento"] = pd.to_datetime(
            df["DataHora_Processamento"],
            errors="coerce"
        )

        df = df.sort_values(
            by="DataHora_Processamento",
            ascending=False
        )

        df = df.drop_duplicates(
            subset=["Codigo"],
            keep="first"
        )

    # Criar status
    def definir_status(row):
        saldo = row["Saldo Total"]
        demanda = row["Demanda"]

        if saldo <= 0:
            return "FALTA"
        elif saldo < demanda:
            return "RISCO"
        elif saldo <= demanda * 1.5:
            return "ATENCAO"
        else:
            return "OK"

    df["Status"] = df.apply(definir_status, axis=1)

    # Sidebar
    st.sidebar.header("Filtros")

    busca_codigo = st.sidebar.text_input("Buscar Código")
    busca_descricao = st.sidebar.text_input("Buscar Descrição")

    status_opcoes = ["TODOS", "OK", "ATENCAO", "RISCO", "FALTA"]
    filtro_sidebar = st.sidebar.selectbox("Filtrar Status", status_opcoes)

    # Aplicar busca
    if busca_codigo:
        df = df[
            df["Codigo"].astype(str).str.contains(
                busca_codigo,
                case=False,
                na=False
            )
        ]

    if busca_descricao:
        df = df[
            df["Descricao"].astype(str).str.contains(
                busca_descricao,
                case=False,
                na=False
            )
        ]

    if filtro_sidebar != "TODOS":
        df = df[df["Status"] == filtro_sidebar]

    # Controle dos botões
    if "filtro_status" not in st.session_state:
        st.session_state["filtro_status"] = "TODOS"

    total_itens = len(df)
    total_ok = len(df[df["Status"] == "OK"])
    total_atencao = len(df[df["Status"] == "ATENCAO"])
    total_risco = len(df[df["Status"] == "RISCO"])
    total_falta = len(df[df["Status"] == "FALTA"])

    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        if st.button("TODOS: " + str(total_itens), use_container_width=True):
            st.session_state["filtro_status"] = "TODOS"

    with col2:
        if st.button("OK: " + str(total_ok), use_container_width=True):
            st.session_state["filtro_status"] = "OK"

    with col3:
        if st.button("ATENCAO: " + str(total_atencao), use_container_width=True):
            st.session_state["filtro_status"] = "ATENCAO"

    with col4:
        if st.button("RISCO: " + str(total_risco), use_container_width=True):
            st.session_state["filtro_status"] = "RISCO"

    with col5:
        if st.button("FALTA: " + str(total_falta), use_container_width=True):
            st.session_state["filtro_status"] = "FALTA"

    # Aplicar filtro dos botões
    if st.session_state["filtro_status"] != "TODOS":
        df_filtrado = df[df["Status"] == st.session_state["filtro_status"]]
    else:
        df_filtrado = df.copy()

    # Funções de cor
    def colorir_status(valor):
        if valor == "FALTA":
            return "background-color: #ff4d4d; color: white; font-weight: bold"
        elif valor == "RISCO":
            return "background-color: #ffd966; color: black; font-weight: bold"
        elif valor == "ATENCAO":
            return "background-color: #f6b26b; color: black; font-weight: bold"
        elif valor == "OK":
            return "background-color: #93c47d; color: black; font-weight: bold"
        return ""

    def colorir_saldo(valor):
        try:
            valor = float(valor)

            if valor < 0:
                return "background-color: #ff4d4d; color: white"
            elif valor == 0:
                return "background-color: #f4cccc; color: black"
            elif valor <= 100:
                return "background-color: #ffe599; color: black"
            else:
                return "background-color: #b6d7a8; color: black"
        except:
            return ""

    st.subheader("Resumo Geral")

    m1, m2, m3, m4 = st.columns(4)

    with m1:
        st.metric("Itens", total_itens)

    with m2:
        st.metric("Itens OK", total_ok)

    with m3:
        st.metric("Itens em Risco", total_risco)

    with m4:
        st.metric("Itens em Falta", total_falta)

    st.subheader("Tabela de Itens")

    tabela_estilizada = (
        df_filtrado.style
        .applymap(colorir_status, subset=["Status"])
        .applymap(colorir_saldo, subset=["Saldo Almox 1", "Saldo Total"])
    )

    st.dataframe(
        tabela_estilizada,
        use_container_width=True,
        height=600
    )

    csv = df_filtrado.to_csv(index=False).encode("utf-8")

    st.download_button(
        label="Baixar Dados Filtrados",
        data=csv,
        file_name="dashboard_filtrado.csv",
        mime="text/csv"
    )
