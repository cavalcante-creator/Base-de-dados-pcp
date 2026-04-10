# =========================
# STATUS AUTOMÁTICO
# =========================
def status(row):
    if row["Saldo vs Demanda"] < 0:
        return "FALTA"
    if row["Demanda Pedido"] >= row["Saldo Almox 3"] * 0.5:
        return "RISCO"
    return "OK"

df["Status Auto"] = df.apply(status, axis=1)

# =========================
# STATUS MANUAL (NOVO)
# =========================
if "Status Manual" not in st.session_state:
    st.session_state["Status Manual"] = {}

df["Status Manual"] = df["Codigo"].map(st.session_state["Status Manual"])

# PRIORIDADE: MANUAL > AUTO
df["Status Final"] = df["Status Manual"].fillna(df["Status Auto"])

# =========================
# CARDS (USANDO STATUS FINAL)
# =========================
col1, col2, col3, col4 = st.columns(4)
col1.metric("Total de Itens", len(df))
col2.metric("Itens em Falta", int((df["Status Final"] == "FALTA").sum()))
col3.metric("Itens em Risco", int((df["Status Final"] == "RISCO").sum()))
col4.metric("Itens OK", int((df["Status Final"] == "OK").sum()))

st.markdown("---")

# =========================
# FILTROS
# =========================
col_filtro1, col_filtro2 = st.columns([1, 2])

with col_filtro1:
    opcoes_status = ["FALTA", "RISCO", "OK"]
    status_selecionado = st.multiselect(
        "Filtrar Status",
        options=opcoes_status,
        default=opcoes_status
    )

with col_filtro2:
    texto_busca = st.text_input("Buscar por código ou descrição")

df_filtrado = df[df["Status Final"].isin(status_selecionado)].copy()

if texto_busca:
    filtro = texto_busca.strip().lower()
    df_filtrado = df_filtrado[
        df_filtrado["Codigo"].astype(str).str.lower().str.contains(filtro, na=False) |
        df_filtrado["Descricao"].astype(str).str.lower().str.contains(filtro, na=False)
    ]

# =========================
# GRÁFICO
# =========================
st.subheader("Visão Geral por Status")
st.bar_chart(df["Status Final"].value_counts())

# =========================
# TABELA EDITÁVEL (NOVO)
# =========================
st.subheader("Tabela de Análise (com ajuste manual)")

df_editado = st.data_editor(
    df_filtrado,
    column_config={
        "Status Manual": st.column_config.SelectboxColumn(
            "Status Manual",
            options=["FALTA", "RISCO", "OK"],
            required=False
        )
    },
    use_container_width=True,
    key="editor"
)

# SALVAR ALTERAÇÕES
for _, row in df_editado.iterrows():
    if pd.notna(row["Status Manual"]):
        st.session_state["Status Manual"][row["Codigo"]] = row["Status Manual"]

# =========================
# DOWNLOAD
# =========================
botao_downloads(df_editado, "dashboard_pcp", "Dashboard_PCP")
