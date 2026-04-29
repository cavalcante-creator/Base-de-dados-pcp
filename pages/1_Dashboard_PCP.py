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
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

*, *::before, *::after { box-sizing: border-box; }

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.main { background: #f0f4f8; }

.block-container {
    padding-top: 1.5rem !important;
    padding-bottom: 2rem !important;
    max-width: 100% !important;
}

/* ---- CARDS DE STATUS ---- */
.cards-row {
    display: flex;
    gap: 14px;
    margin-bottom: 20px;
    flex-wrap: wrap;
}

.status-card {
    flex: 1;
    min-width: 150px;
    border-radius: 16px;
    padding: 18px 20px;
    display: flex;
    flex-direction: column;
    gap: 4px;
    cursor: pointer;
    transition: transform .15s, box-shadow .15s;
    box-shadow: 0 2px 10px rgba(0,0,0,.10);
    border: 2px solid transparent;
    color: white;
}
.status-card:hover { transform: translateY(-2px); box-shadow: 0 6px 20px rgba(0,0,0,.15); }
.card-total   { background: linear-gradient(135deg, #1d4ed8, #3b82f6); }
.card-falta   { background: linear-gradient(135deg, #dc2626, #ef4444); }
.card-risco   { background: linear-gradient(135deg, #d97706, #f59e0b); }
.card-ok      { background: linear-gradient(135deg, #16a34a, #22c55e); }
.card-selected { border-color: rgba(255,255,255,.7); transform: scale(1.03); }

.card-label { font-size: 12px; font-weight: 600; letter-spacing: .08em; text-transform: uppercase; opacity: .85; }
.card-num   { font-size: 36px; font-weight: 700; line-height: 1; }
.card-sub   { font-size: 11px; opacity: .75; }

/* ---- BADGE STATUS ---- */
.badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 999px;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: .05em;
}
.badge-falta { background:#fee2e2; color:#991b1b; }
.badge-risco { background:#fef3c7; color:#92400e; }
.badge-ok    { background:#dcfce7; color:#166534; }

/* ---- SUGESTÃO BOX ---- */
.sugestao-box {
    background: linear-gradient(135deg, #1e3a5f 0%, #1d4ed8 100%);
    border-radius: 18px;
    padding: 22px 26px;
    margin-bottom: 18px;
    color: white;
    box-shadow: 0 6px 20px rgba(29,78,216,.3);
}
.sugestao-title {
    font-size: 15px;
    font-weight: 700;
    letter-spacing: .05em;
    margin-bottom: 14px;
    opacity: .9;
}
.sugestao-item {
    background: rgba(255,255,255,.12);
    border-radius: 10px;
    padding: 12px 16px;
    margin-bottom: 8px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    border: 1px solid rgba(255,255,255,.1);
}
.sugestao-item-cod { font-weight: 600; font-size: 13px; }
.sugestao-item-desc { font-size: 12px; opacity: .75; }
.sugestao-qtd {
    background: rgba(255,255,255,.2);
    border-radius: 8px;
    padding: 4px 10px;
    font-weight: 700;
    font-size: 13px;
    white-space: nowrap;
}
.urgente { border-left: 4px solid #ef4444 !important; }
.atencao { border-left: 4px solid #f59e0b !important; }

/* ---- ALERTA BOX ---- */
.alerta-seg {
    background: #fffbeb;
    border: 1px solid #fde68a;
    border-radius: 14px;
    padding: 16px 20px;
    margin-top: 16px;
}

/* ---- FILTROS BLOCO ---- */
.filtro-bloco {
    background: white;
    border-radius: 14px;
    padding: 16px 20px;
    margin-bottom: 16px;
    border: 1px solid #e2e8f0;
    box-shadow: 0 1px 6px rgba(0,0,0,.05);
}

/* ---- DATAFRAME ---- */
div[data-testid="stDataFrame"] {
    background: white !important;
    border-radius: 16px !important;
    border: 1px solid #e2e8f0 !important;
    box-shadow: 0 2px 12px rgba(0,0,0,.06) !important;
    overflow: hidden;
}

/* ---- INPUTS ---- */
div[data-testid="stTextInput"] > div > div > input {
    border-radius: 10px !important;
    border: 1.5px solid #cbd5e1 !important;
    padding: 8px 14px !important;
}
div[data-testid="stTextInput"] > div > div > input:focus {
    border-color: #3b82f6 !important;
    box-shadow: 0 0 0 3px rgba(59,130,246,.15) !important;
}

/* ---- BOTÕES STREAMLIT ---- */
div.stButton > button {
    border-radius: 10px !important;
    font-weight: 600 !important;
    border: 1.5px solid #e2e8f0 !important;
    background: white !important;
    color: #1e293b !important;
    transition: all .15s !important;
}
div.stButton > button:hover {
    background: #f1f5f9 !important;
    border-color: #94a3b8 !important;
}

/* ---- HEADER ---- */
.pcp-header {
    background: linear-gradient(135deg, #0f172a 0%, #1d4ed8 100%);
    border-radius: 18px;
    padding: 20px 28px;
    margin-bottom: 20px;
    color: white;
    display: flex;
    align-items: center;
    justify-content: space-between;
    box-shadow: 0 6px 24px rgba(15,23,42,.3);
}
.pcp-header-title { font-size: 22px; font-weight: 700; }
.pcp-header-sub { font-size: 13px; opacity: .7; margin-top: 3px; }
.pcp-header-tag {
    background: rgba(255,255,255,.12);
    border-radius: 8px;
    padding: 6px 14px;
    font-size: 12px;
    font-weight: 600;
    border: 1px solid rgba(255,255,255,.2);
}

/* ---- MULTISELECT ---- */
span[data-baseweb="tag"] {
    background: #dbeafe !important;
    color: #1d4ed8 !important;
    border-radius: 6px !important;
}
</style>
""", unsafe_allow_html=True)

# ========================= FUNÇÕES =========================
fuso = pytz.timezone("America/Sao_Paulo")

def agora():
    return datetime.now(fuso)

def exportar_excel_formatado(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Dados")
        ws = writer.book["Dados"]
        header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
        header_font = Font(color="FFFFFF", bold=True)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for idx, coluna in enumerate(df.columns, start=1):
            letra = get_column_letter(idx)
            ws.column_dimensions[letra].width = min(
                max(len(str(coluna)) + 4,
                    max((len(str(v)) for v in df[coluna] if pd.notna(v)), default=4) + 2),
                40
            )
        if "Status" in df.columns:
            idx_s = list(df.columns).index("Status") + 1
            for cell in ws[get_column_letter(idx_s)][1:]:
                v = str(cell.value or "")
                if "FALTA" in v:
                    cell.fill = PatternFill(fill_type="solid", fgColor="FFC7CE")
                elif "RISCO" in v:
                    cell.fill = PatternFill(fill_type="solid", fgColor="FFF2CC")
                elif "OK" in v:
                    cell.fill = PatternFill(fill_type="solid", fgColor="C6E0B4")
    output.seek(0)
    return output.getvalue()

# ========================= DADOS =========================
try:
    saldo = pd.read_csv("saldo.csv")
    perfil = pd.read_csv("perfil.csv")
    previsao = pd.read_csv("previsao.csv")
except Exception:
    st.warning("⚠️ Faça upload dos dados na página principal primeiro.")
    st.stop()

# ========================= ORDENS EM ABERTO =========================
tem_ordens = False
ordens_abertas = pd.DataFrame()

try:
    ordens_raw = pd.read_csv("ordens.csv")
    colunas_upper = {c: c.upper().strip() for c in ordens_raw.columns}
    col_cod_ordens = col_qtd_ordens = col_status_ordens = None

    for orig, upper in colunas_upper.items():
        if any(k in upper for k in ["COD", "ITEM", "PRODUTO", "CODIGO", "CÓDIGO"]) and col_cod_ordens is None:
            col_cod_ordens = orig
        if any(k in upper for k in ["QTDE", "QTD", "QUANTIDADE", "SALDO", "SALDO ORDEM"]) and col_qtd_ordens is None:
            col_qtd_ordens = orig
        if any(k in upper for k in ["STATUS", "SITUAÇÃO", "SITUACAO", "SIT"]) and col_status_ordens is None:
            col_status_ordens = orig

    if col_cod_ordens and col_qtd_ordens:
        cols = [col_cod_ordens, col_qtd_ordens] + ([col_status_ordens] if col_status_ordens else [])
        df_ord = ordens_raw[cols].copy()
        df_ord.columns = ["Codigo", "Qtde Ordem"] + (["Status Ordem"] if col_status_ordens else [])

        if "Status Ordem" in df_ord.columns:
            palavras_aberto = ["ABERTO", "ABERTA", "EM ABERTO", "LIBERADA", "LIBERADO", "PENDENTE", "A PRODUZIR"]
            mask = df_ord["Status Ordem"].astype(str).str.upper().str.strip().isin(palavras_aberto)
            if mask.sum() > 0:
                df_ord = df_ord[mask]

        df_ord["Qtde Ordem"] = (
            df_ord["Qtde Ordem"].astype(str)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        df_ord["Qtde Ordem"] = pd.to_numeric(df_ord["Qtde Ordem"], errors="coerce").fillna(0)
        ordens_abertas = df_ord.groupby("Codigo")["Qtde Ordem"].sum().reset_index()
        ordens_abertas.columns = ["Codigo", "Qtde Ordens Abertas"]
        tem_ordens = True
except Exception:
    pass

try:
    parametros = pd.read_csv("parametros.csv")
    parametros = parametros.rename(columns={"COD ITEM": "Codigo", "ESTQ SEG": "Estq Seg"})
    parametros = parametros[["Codigo", "Estq Seg"]].copy()
    parametros["Estq Seg"] = pd.to_numeric(parametros["Estq Seg"], errors="coerce").fillna(0)
    tem_parametros = True
except Exception:
    tem_parametros = False

saldo = saldo.sort_values(by=["Data Processamento", "Hora Processamento"], ascending=False).drop_duplicates("Codigo")
previsao = previsao.sort_values(by=["Data Processamento", "Hora Processamento"], ascending=False).drop_duplicates("COD")

base = previsao[["COD", "PRODUTO"]].copy()
base.columns = ["Codigo", "Descricao"]

# ========================= ALMOXARIFADOS =========================
colunas_almox = [c for c in saldo.columns if "SALDO" in c.upper() and "ALMOX" in c.upper()]
opcoes_almox = (["Saldo Total"] if "Saldo Total" in saldo.columns else []) + colunas_almox

if not opcoes_almox:
    st.error("Nenhuma coluna de saldo encontrada. Reprocesse o PDF de Saldo.")
    st.stop()

for col in opcoes_almox:
    if col not in saldo.columns:
        saldo[col] = 0

saldo_base = saldo[["Codigo"] + opcoes_almox]

perfil["Quantidade"] = (
    perfil["Quantidade"].astype(str)
    .str.replace(".", "", regex=False)
    .str.replace(",", ".", regex=False)
    .astype(float)
)

dc = perfil[perfil["Tipo"] == "DC"].groupby("Item")["Quantidade"].sum().reset_index()
dc.columns = ["Codigo", "Demanda Pedido"]

tipos_op = ["OP", "ORDEM", "LIBERADA"]
op = perfil[perfil["Tipo"].isin(tipos_op)].groupby("Item")["Quantidade"].sum().reset_index()
op.columns = ["Codigo", "Qtde Pendente OP"]

df_full = base.merge(saldo_base, on="Codigo", how="left")
df_full = df_full.merge(dc, on="Codigo", how="left")
df_full = df_full.merge(op, on="Codigo", how="left")

if tem_ordens:
    df_full = df_full.merge(ordens_abertas, on="Codigo", how="left")

if tem_parametros:
    df_full = df_full.merge(parametros, on="Codigo", how="left")
    df_full["Estq Seg"] = df_full["Estq Seg"].fillna(0)
else:
    df_full["Estq Seg"] = 0

df_full = df_full.fillna(0)

# ========================= HEADER =========================
now_str = agora().strftime("%d/%m/%Y %H:%M")
st.markdown(f"""
<div class="pcp-header">
    <div>
        <div class="pcp-header-title">📊 Dashboard PCP</div>
        <div class="pcp-header-sub">Planejamento e Controle da Produção</div>
    </div>
    <div class="pcp-header-tag">🕐 {now_str}</div>
</div>
""", unsafe_allow_html=True)

# ========================= CONFIGURAÇÕES =========================
with st.expander("⚙️ Configuração dos Almoxarifados e Filtros", expanded=False):
    col_cfg1, col_cfg2, col_cfg3 = st.columns(3)

    with col_cfg1:
        idx_demanda = opcoes_almox.index("Saldo Almox 3") if "Saldo Almox 3" in opcoes_almox else 0
        almox_demanda = st.selectbox(
            "Almox para **Demanda / Pedidos**",
            options=opcoes_almox,
            index=idx_demanda,
            help="Saldo deste almox será comparado com a demanda de pedidos"
        )

    with col_cfg2:
        idx_seg = opcoes_almox.index("Saldo Almox 30") if "Saldo Almox 30" in opcoes_almox else min(1, len(opcoes_almox) - 1)
        almox_estq_seg = st.selectbox(
            "Almox para **Estoque de Segurança**",
            options=opcoes_almox,
            index=idx_seg,
            help="Saldo deste almox será comparado com o estoque de segurança"
        )

    with col_cfg3:
        if tem_ordens:
            considerar_ordens = st.checkbox(
                "📋 Incluir Ordens em Aberto no Saldo Real",
                value=True,
                help="Soma as ordens de produção em aberto ao saldo disponível"
            )
        else:
            st.info("ℹ️ Ordens não carregadas")
            considerar_ordens = False

# ========================= CÁLCULOS =========================
df_full["Saldo vs Demanda"] = df_full[almox_demanda] - df_full["Demanda Pedido"]

saldo_real = df_full[almox_demanda].copy()
if considerar_ordens and tem_ordens:
    saldo_real = saldo_real + df_full["Qtde Ordens Abertas"]

df_full["Saldo Real"] = saldo_real - df_full["Demanda Pedido"] - df_full["Qtde Pendente OP"]
df_full["Abaixo Estq Seg"] = df_full[almox_estq_seg] < df_full["Estq Seg"]

def calcular_status(row):
    if row["Saldo Real"] < 0:
        return "FALTA"
    if row["Abaixo Estq Seg"]:
        return "RISCO"
    if row["Demanda Pedido"] + row["Qtde Pendente OP"] >= row[almox_estq_seg] * 0.5:
        return "RISCO"
    return "OK"

df_full["Status"] = df_full.apply(calcular_status, axis=1)

# ========================= FILTRO DE ITENS (EXCLUSÃO) =========================
st.markdown('<div class="filtro-bloco">', unsafe_allow_html=True)
col_f1, col_f2 = st.columns([2, 3])

with col_f1:
    busca = st.text_input("🔍 Pesquisar por código ou descrição", placeholder="Ex: ABC123 ou nome do produto...")

with col_f2:
    todos_codigos = sorted(df_full["Codigo"].astype(str).unique().tolist())
    itens_excluir = st.multiselect(
        "🚫 Ocultar itens específicos",
        options=todos_codigos,
        placeholder="Selecione os itens que não devem aparecer..."
    )

st.markdown('</div>', unsafe_allow_html=True)

# ========================= FILTRO ESTOQUE DE SEGURANÇA =========================
if tem_parametros:
    with st.expander("🔒 Filtros de Estoque de Segurança", expanded=False):
        col_es1, col_es2 = st.columns(2)
        with col_es1:
            busca_estq_seg = st.text_input(
                "🔍 Pesquisar no Estoque de Segurança",
                placeholder="Código ou descrição...",
                key="busca_estq"
            )
        with col_es2:
            filtro_apenas_abaixo = st.checkbox("Mostrar apenas itens abaixo do estoque de segurança", value=False)

# ========================= FILTRO STATUS (SESSION STATE) =========================
if "filtro_status" not in st.session_state:
    st.session_state.filtro_status = "TODOS"

# ========================= CARDS =========================
n_total = len(df_full)
n_falta = (df_full["Status"] == "FALTA").sum()
n_risco = (df_full["Status"] == "RISCO").sum()
n_ok = (df_full["Status"] == "OK").sum()

c1, c2, c3, c4 = st.columns(4)

cards = [
    (c1, "TODOS", "card-total", "📦 Total", n_total, "itens na base"),
    (c2, "FALTA", "card-falta", "🔴 Falta", n_falta, "precisam produção urgente"),
    (c3, "RISCO", "card-risco", "🟡 Risco", n_risco, "abaixo do estoque seguro"),
    (c4, "OK",    "card-ok",   "🟢 OK",    n_ok,    "dentro do esperado"),
]

for col, key, css, label, num, sub in cards:
    sel = "card-selected" if st.session_state.filtro_status == key else ""
    with col:
        if st.button(f"{label}  •  {num}", key=f"btn_{key}", use_container_width=True):
            st.session_state.filtro_status = key
        st.markdown(f"""
        <div class="status-card {css} {sel}" style="margin-top:-8px;">
            <div class="card-label">{label}</div>
            <div class="card-num">{num}</div>
            <div class="card-sub">{sub}</div>
        </div>
        """, unsafe_allow_html=True)

st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

# ========================= APLICAR FILTROS =========================
df_filtrado = df_full.copy()

# Exclusão de itens
if itens_excluir:
    df_filtrado = df_filtrado[~df_filtrado["Codigo"].astype(str).isin(itens_excluir)]

# Filtro de status
if st.session_state.filtro_status != "TODOS":
    df_filtrado = df_filtrado[df_filtrado["Status"] == st.session_state.filtro_status]

# Busca
if busca:
    df_filtrado = df_filtrado[
        df_filtrado["Codigo"].astype(str).str.contains(busca, case=False, na=False) |
        df_filtrado["Descricao"].astype(str).str.contains(busca, case=False, na=False)
    ]

# ========================= LEGENDA =========================
col_leg1, col_leg2 = st.columns(2)
col_leg1.caption(f"📦 Demanda/Pedidos usando: `{almox_demanda}`")
col_leg2.caption(f"🔒 Estq. Segurança usando: `{almox_estq_seg}`" + (" + 📋 Ordens em Aberto" if considerar_ordens and tem_ordens else ""))

if itens_excluir:
    st.caption(f"🚫 {len(itens_excluir)} item(ns) oculto(s) da visualização")

# ========================= COLUNAS EXIBIDAS =========================
colunas_exibir = ["Codigo", "Descricao", almox_demanda]
if almox_estq_seg != almox_demanda:
    colunas_exibir.append(almox_estq_seg)
colunas_exibir += ["Demanda Pedido", "Qtde Pendente OP"]
if tem_ordens:
    colunas_exibir.append("Qtde Ordens Abertas")
colunas_exibir += ["Saldo vs Demanda", "Saldo Real"]
if tem_parametros:
    colunas_exibir.append("Estq Seg")
colunas_exibir.append("Status")
colunas_exibir = [c for c in colunas_exibir if c in df_filtrado.columns]

# ========================= ESTILO POR LINHA =========================
def estilo_linha(row):
    if row["Status"] == "FALTA":
        return ["background-color:#fee2e2; color:#7f1d1d"] * len(row)
    if row["Status"] == "RISCO":
        return ["background-color:#fef9c3; color:#78350f"] * len(row)
    if row["Status"] == "OK":
        return ["background-color:#f0fdf4; color:#14532d"] * len(row)
    return [""] * len(row)

# ========================= TABELA PRINCIPAL =========================
st.markdown(f"**{len(df_filtrado)} item(ns) exibido(s)**")
st.dataframe(
    df_filtrado[colunas_exibir].style.apply(estilo_linha, axis=1),
    use_container_width=True,
    height=420
)

# ========================= SUGESTÃO DE PRODUÇÃO =========================
st.markdown("---")
st.markdown("### 🏭 Sugestão de Produção")

df_sugestao = df_full[df_full["Status"].isin(["FALTA", "RISCO"])].copy()

if not df_sugestao.empty:
    df_sugestao["Qtde Sugerida"] = df_sugestao.apply(
        lambda row: max(
            row["Demanda Pedido"] + row["Qtde Pendente OP"] - row[almox_demanda] + row["Estq Seg"],
            0
        ),
        axis=1
    )
    df_sugestao = df_sugestao[df_sugestao["Qtde Sugerida"] > 0].sort_values(
        by=["Status", "Qtde Sugerida"], ascending=[True, False]
    )

    n_urgente = (df_sugestao["Status"] == "FALTA").sum()
    n_atencao = (df_sugestao["Status"] == "RISCO").sum()

    col_s1, col_s2, col_s3 = st.columns([3, 1, 1])
    with col_s1:
        st.markdown(f"""
        <div class="sugestao-box">
            <div class="sugestao-title">📋 {len(df_sugestao)} ITEM(NS) PRECISAM SER PRODUZIDOS</div>
        """, unsafe_allow_html=True)

        for _, row in df_sugestao.head(10).iterrows():
            cls = "urgente" if row["Status"] == "FALTA" else "atencao"
            emoji = "🔴" if row["Status"] == "FALTA" else "🟡"
            st.markdown(f"""
            <div class="sugestao-item {cls}">
                <div>
                    <div class="sugestao-item-cod">{emoji} {row['Codigo']}</div>
                    <div class="sugestao-item-desc">{str(row['Descricao'])[:50]}</div>
                </div>
                <div class="sugestao-qtd">Prod: {row['Qtde Sugerida']:,.0f}</div>
            </div>
            """, unsafe_allow_html=True)

        if len(df_sugestao) > 10:
            st.markdown(f'<div style="color:rgba(255,255,255,.6); font-size:12px; margin-top:8px;">+ {len(df_sugestao)-10} outros itens (veja o Excel)</div>', unsafe_allow_html=True)

        st.markdown("</div>", unsafe_allow_html=True)

    with col_s2:
        st.metric("🔴 Urgente", n_urgente, help="Status FALTA — saldo negativo")
    with col_s3:
        st.metric("🟡 Atenção", n_atencao, help="Status RISCO — abaixo do estoque de segurança")

    # Download sugestão
    df_down_sug = df_sugestao[["Codigo", "Descricao", almox_demanda, "Demanda Pedido", "Qtde Pendente OP", "Estq Seg", "Qtde Sugerida", "Status"]].copy()
    st.download_button(
        "📥 Baixar Sugestão de Produção (Excel)",
        exportar_excel_formatado(df_down_sug),
        "sugestao_producao.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_sugestao"
    )
else:
    st.success("✅ Nenhum item necessita produção no momento. Todos os saldos estão adequados.")

# ========================= ALERTA ESTOQUE DE SEGURANÇA =========================
st.markdown("---")
if tem_parametros:
    df_abaixo = df_full[df_full["Abaixo Estq Seg"]].copy()

    # Aplicar busca do estoque de segurança
    if tem_parametros:
        try:
            if busca_estq_seg:
                df_abaixo = df_abaixo[
                    df_abaixo["Codigo"].astype(str).str.contains(busca_estq_seg, case=False, na=False) |
                    df_abaixo["Descricao"].astype(str).str.contains(busca_estq_seg, case=False, na=False)
                ]
        except Exception:
            pass

    if not df_abaixo.empty:
        with st.expander(f"⚠️ {len(df_abaixo)} item(ns) abaixo do Estoque de Segurança ({almox_estq_seg})", expanded=True):
            # Busca no expander
            col_alerta_f1, col_alerta_f2 = st.columns([3, 1])
            with col_alerta_f1:
                busca_alerta = st.text_input(
                    "🔍 Pesquisar no estoque de segurança",
                    placeholder="Código ou descrição...",
                    key="busca_alerta_inline"
                )
            with col_alerta_f2:
                ordem_alerta = st.selectbox("Ordenar por", ["Diferença (maior)", "Código", "Status"], key="ord_alerta")

            df_alerta = df_abaixo.copy()
            cols_alerta = ["Codigo", "Descricao", almox_estq_seg, "Estq Seg", "Status"]
            cols_alerta = [c for c in cols_alerta if c in df_alerta.columns]
            df_alerta = df_alerta[cols_alerta].copy()
            df_alerta["Diferença"] = df_alerta[almox_estq_seg] - df_alerta["Estq Seg"]

            if busca_alerta:
                df_alerta = df_alerta[
                    df_alerta["Codigo"].astype(str).str.contains(busca_alerta, case=False, na=False) |
                    df_alerta["Descricao"].astype(str).str.contains(busca_alerta, case=False, na=False)
                ]

            if ordem_alerta == "Diferença (maior)":
                df_alerta = df_alerta.sort_values("Diferença")
            elif ordem_alerta == "Código":
                df_alerta = df_alerta.sort_values("Codigo")
            elif ordem_alerta == "Status":
                df_alerta = df_alerta.sort_values("Status")

            st.dataframe(df_alerta.style.apply(estilo_linha, axis=1), use_container_width=True)

            st.download_button(
                "📥 Baixar Alertas Estoque Segurança",
                exportar_excel_formatado(df_alerta),
                "alertas_estoque_seg.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_alerta_estq"
            )
    else:
        st.success("✅ Todos os itens estão acima do Estoque de Segurança.")
else:
    st.info("ℹ️ Importe o arquivo de Parâmetros para comparar com o Estoque de Segurança.")

# ========================= DOWNLOAD GERAL =========================
st.markdown("---")
col_d1, col_d2 = st.columns(2)
with col_d1:
    st.download_button(
        "📥 Baixar Tabela Atual (CSV)",
        df_filtrado[colunas_exibir].to_csv(index=False).encode("utf-8"),
        "pcp_dashboard.csv",
        mime="text/csv",
        key="dl_csv_main"
    )
with col_d2:
    st.download_button(
        "📥 Baixar Tabela Atual (Excel)",
        exportar_excel_formatado(df_filtrado[colunas_exibir]),
        "pcp_dashboard.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_excel_main"
    )
