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
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
*, *::before, *::after { box-sizing: border-box; }
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.main { background: #f0f4f8; }
.block-container { padding-top: 1rem !important; padding-bottom: 2rem !important; max-width: 100% !important; }
.pcp-header { background: linear-gradient(135deg, #0f172a 0%, #1d4ed8 100%); border-radius: 18px; padding: 20px 28px; margin-bottom: 20px; color: white; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 6px 24px rgba(15,23,42,.3); }
.pcp-header-title { font-size: 22px; font-weight: 800; }
.pcp-header-sub { font-size: 13px; opacity: .7; margin-top: 3px; }
.pcp-header-tag { background: rgba(255,255,255,.12); border-radius: 8px; padding: 6px 14px; font-size: 12px; font-weight: 600; border: 1px solid rgba(255,255,255,.2); }
.bloco-card { background: white; border-radius: 16px; border: 1px solid #e2e8f0; box-shadow: 0 2px 12px rgba(0,0,0,.06); padding: 20px 24px; margin-bottom: 16px; }
.bloco-title { font-size: 15px; font-weight: 700; color: #0f172a; margin-bottom: 14px; }
.status-card { border-radius: 14px; padding: 16px 18px; color: white; box-shadow: 0 2px 10px rgba(0,0,0,.12); }
.card-total { background: linear-gradient(135deg, #1d4ed8, #3b82f6); }
.card-falta { background: linear-gradient(135deg, #dc2626, #ef4444); }
.card-risco { background: linear-gradient(135deg, #d97706, #f59e0b); }
.card-ok { background: linear-gradient(135deg, #16a34a, #22c55e); }
.card-label { font-size: 11px; font-weight: 600; letter-spacing: .08em; text-transform: uppercase; opacity: .85; }
.card-num { font-size: 34px; font-weight: 800; line-height: 1.1; }
.card-sub { font-size: 11px; opacity: .75; margin-top: 2px; }
.sugestao-box { background: linear-gradient(135deg, #1e3a5f 0%, #1d4ed8 100%); border-radius: 16px; padding: 20px 24px; margin-bottom: 16px; color: white; box-shadow: 0 6px 20px rgba(29,78,216,.3); }
.sugestao-item { background: rgba(255,255,255,.12); border-radius: 10px; padding: 11px 14px; margin-bottom: 7px; display: flex; justify-content: space-between; align-items: center; border: 1px solid rgba(255,255,255,.1); }
.urgente { border-left: 4px solid #ef4444 !important; }
.atencao { border-left: 4px solid #f59e0b !important; }
div[data-testid="stDataFrame"] { background: white !important; border-radius: 16px !important; border: 1px solid #e2e8f0 !important; box-shadow: 0 2px 12px rgba(0,0,0,.06) !important; overflow: hidden; }
div.stButton > button { border-radius: 10px !important; font-weight: 600 !important; border: 1.5px solid #e2e8f0 !important; background: white !important; color: #1e293b !important; transition: all .15s !important; }
div.stButton > button:hover { background: #f1f5f9 !important; border-color: #94a3b8 !important; }
span[data-baseweb="tag"] { background: #dbeafe !important; color: #1d4ed8 !important; border-radius: 6px !important; }
.secao-titulo { font-size: 17px; font-weight: 700; color: #0f172a; margin: 20px 0 10px 0; padding-bottom: 8px; border-bottom: 2px solid #e2e8f0; }
.info-pill { display: inline-flex; align-items: center; gap: 6px; background: #eff6ff; border: 1px solid #bfdbfe; border-radius: 999px; padding: 4px 12px; font-size: 12px; color: #1d4ed8; font-weight: 500; margin-right: 6px; margin-bottom: 6px; }
</style>
""", unsafe_allow_html=True)

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
            ws.column_dimensions[letra].width = min(max(len(str(coluna)) + 4, max((len(str(v)) for v in df[coluna] if pd.notna(v)), default=4) + 2), 40)
        if "Status" in df.columns:
            idx_s = list(df.columns).index("Status") + 1
            for cell in ws[get_column_letter(idx_s)][1:]:
                v = str(cell.value or "")
                if "FALTA" in v: cell.fill = PatternFill(fill_type="solid", fgColor="FFC7CE")
                elif "RISCO" in v: cell.fill = PatternFill(fill_type="solid", fgColor="FFF2CC")
                elif "OK" in v: cell.fill = PatternFill(fill_type="solid", fgColor="C6E0B4")
    output.seek(0)
    return output.getvalue()

def estilo_linha(row):
    if row["Status"] == "FALTA": return ["background-color:#fee2e2; color:#7f1d1d"] * len(row)
    if row["Status"] == "RISCO": return ["background-color:#fef9c3; color:#78350f"] * len(row)
    if row["Status"] == "OK": return ["background-color:#f0fdf4; color:#14532d"] * len(row)
    return [""] * len(row)

try:
    saldo = pd.read_csv("saldo.csv")
    perfil = pd.read_csv("perfil.csv")
    previsao = pd.read_csv("previsao.csv")
except Exception:
    st.warning("⚠️ Faça upload dos dados na página principal primeiro.")
    st.stop()

tem_ordens = False
ordens_abertas = pd.DataFrame()
try:
    ordens_raw = pd.read_csv("ordens.csv")
    colunas_upper = {c: c.upper().strip() for c in ordens_raw.columns}
    col_cod_ordens = col_qtd_ordens = col_status_ordens = None
    for orig, upper in colunas_upper.items():
        if any(k in upper for k in ["COD", "ITEM", "PRODUTO", "CODIGO", "CÓDIGO"]) and col_cod_ordens is None: col_cod_ordens = orig
        if any(k in upper for k in ["QTDE", "QTD", "QUANTIDADE", "SALDO", "SALDO ORDEM"]) and col_qtd_ordens is None: col_qtd_ordens = orig
        if any(k in upper for k in ["STATUS", "SITUAÇÃO", "SITUACAO", "SIT"]) and col_status_ordens is None: col_status_ordens = orig
    if col_cod_ordens and col_qtd_ordens:
        cols = [col_cod_ordens, col_qtd_ordens] + ([col_status_ordens] if col_status_ordens else [])
        df_ord = ordens_raw[cols].copy()
        df_ord.columns = ["Codigo", "Qtde Ordem"] + (["Status Ordem"] if col_status_ordens else [])
        if "Status Ordem" in df_ord.columns:
            palavras_aberto = ["ABERTO", "ABERTA", "EM ABERTO", "LIBERADA", "LIBERADO", "PENDENTE", "A PRODUZIR"]
            mask = df_ord["Status Ordem"].astype(str).str.upper().str.strip().isin(palavras_aberto)
            if mask.sum() > 0: df_ord = df_ord[mask]
        df_ord["Qtde Ordem"] = pd.to_numeric(df_ord["Qtde Ordem"].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False), errors="coerce").fillna(0)
        ordens_abertas = df_ord.groupby("Codigo")["Qtde Ordem"].sum().reset_index()
        ordens_abertas.columns = ["Codigo", "Qtde Ordens Abertas"]
        tem_ordens = True
except Exception:
    pass

tem_parametros = False
try:
    parametros = pd.read_csv("parametros.csv")
    parametros = parametros.rename(columns={"COD ITEM": "Codigo", "ESTQ SEG": "Estq Seg"})
    parametros = parametros[["Codigo", "Estq Seg"]].copy()
    parametros["Estq Seg"] = pd.to_numeric(parametros["Estq Seg"], errors="coerce").fillna(0)
    tem_parametros = True
except Exception:
    pass

saldo = saldo.sort_values(by=["Data Processamento", "Hora Processamento"], ascending=False).drop_duplicates("Codigo")
previsao = previsao.sort_values(by=["Data Processamento", "Hora Processamento"], ascending=False).drop_duplicates("COD")
base = previsao[["COD", "PRODUTO"]].copy()
base.columns = ["Codigo", "Descricao"]

colunas_almox = [c for c in saldo.columns if "SALDO" in c.upper() and "ALMOX" in c.upper()]
opcoes_almox = (["Saldo Total"] if "Saldo Total" in saldo.columns else []) + colunas_almox
if not opcoes_almox:
    st.error("Nenhuma coluna de saldo encontrada. Reprocesse o PDF de Saldo.")
    st.stop()
for col in opcoes_almox:
    if col not in saldo.columns: saldo[col] = 0

saldo_base = saldo[["Codigo"] + opcoes_almox]

perfil["Quantidade"] = pd.to_numeric(perfil["Quantidade"].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False), errors="coerce").fillna(0)

dc = perfil[perfil["Tipo"] == "DC"].groupby("Item")["Quantidade"].sum().reset_index()
dc.columns = ["Codigo", "Demanda Pedido"]
op = perfil[perfil["Tipo"].isin(["OP", "ORDEM", "LIBERADA"])].groupby("Item")["Quantidade"].sum().reset_index()
op.columns = ["Codigo", "Qtde Pendente OP"]

df_full = base.merge(saldo_base, on="Codigo", how="left")
df_full = df_full.merge(dc, on="Codigo", how="left")
df_full = df_full.merge(op, on="Codigo", how="left")
if tem_ordens: df_full = df_full.merge(ordens_abertas, on="Codigo", how="left")
if tem_parametros:
    df_full = df_full.merge(parametros, on="Codigo", how="left")
    df_full["Estq Seg"] = df_full["Estq Seg"].fillna(0)
else:
    df_full["Estq Seg"] = 0
df_full = df_full.fillna(0)

if "bloco_ativo" not in st.session_state: st.session_state.bloco_ativo = "resumo"
if "filtro_status" not in st.session_state: st.session_state.filtro_status = "TODOS"

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

blocos = [("resumo","📊 Resumo"), ("filtros","🔍 Filtros"), ("saldos","💰 Saldos"), ("estoques","📦 Estoques"), ("demandas","📋 Demandas"), ("sugestao","🏭 Sugestão")]
cols_nav = st.columns(len(blocos))
for i, (key, label) in enumerate(blocos):
    with cols_nav[i]:
        if st.button(label, key=f"nav_{key}", use_container_width=True):
            st.session_state.bloco_ativo = key

bloco = st.session_state.bloco_ativo
st.markdown(" ".join([f'<span style="background:{"#1d4ed8" if k==bloco else "#e2e8f0"};color:{"white" if k==bloco else "#64748b"};padding:4px 14px;border-radius:999px;font-size:12px;font-weight:600;">{l}</span>' for k,l in blocos]), unsafe_allow_html=True)
st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

# ============================================================
# RESUMO
# ============================================================
if bloco == "resumo":
    with st.expander("⚙️ Configurar Almoxarifados", expanded=False):
        c1, c2, c3 = st.columns(3)
        with c1:
            idx_d = opcoes_almox.index("Saldo Almox 3") if "Saldo Almox 3" in opcoes_almox else 0
            almox_demanda_res = st.selectbox("Almox — Demanda/Pedidos", opcoes_almox, index=idx_d, key="res_almox_dem")
        with c2:
            idx_s = opcoes_almox.index("Saldo Almox 30") if "Saldo Almox 30" in opcoes_almox else min(1, len(opcoes_almox)-1)
            almox_seg_res = st.selectbox("Almox — Estoque de Segurança", opcoes_almox, index=idx_s, key="res_almox_seg")
        with c3:
            usar_ordens_res = st.checkbox("Incluir Ordens em Aberto", value=True, key="res_ordens") if tem_ordens else False

    saldo_real_res = df_full[almox_demanda_res].copy()
    if usar_ordens_res and tem_ordens: saldo_real_res = saldo_real_res + df_full["Qtde Ordens Abertas"]
    df_res = df_full.copy()
    df_res["Saldo Real"] = saldo_real_res - df_res["Demanda Pedido"] - df_res["Qtde Pendente OP"]
    df_res["Abaixo Estq Seg"] = df_res[almox_seg_res] < df_res["Estq Seg"]

    def calc_status_res(row):
        if row["Saldo Real"] < 0: return "FALTA"
        if row["Abaixo Estq Seg"]: return "RISCO"
        if row["Demanda Pedido"] + row["Qtde Pendente OP"] >= row[almox_seg_res] * 0.5: return "RISCO"
        return "OK"

    df_res["Status"] = df_res.apply(calc_status_res, axis=1)
    n_total = len(df_res); n_falta = (df_res["Status"]=="FALTA").sum(); n_risco = (df_res["Status"]=="RISCO").sum(); n_ok = (df_res["Status"]=="OK").sum()

    c1, c2, c3, c4 = st.columns(4)
    for col, key, css, label, num, sub in [(c1,"TODOS","card-total","📦 Total",n_total,"itens na base"),(c2,"FALTA","card-falta","🔴 Falta",n_falta,"produção urgente"),(c3,"RISCO","card-risco","🟡 Risco",n_risco,"abaixo est. seg."),(c4,"OK","card-ok","🟢 OK",n_ok,"dentro do esperado")]:
        with col:
            if st.button(f"{label}  •  {num}", key=f"btn_res_{key}", use_container_width=True):
                st.session_state.filtro_status = key; st.session_state.bloco_ativo = "filtros"; st.rerun()
            st.markdown(f'<div class="status-card {css}" style="margin-top:-8px;"><div class="card-label">{label}</div><div class="card-num">{num}</div><div class="card-sub">{sub}</div></div>', unsafe_allow_html=True)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    col_ex = [c for c in ["Codigo","Descricao",almox_demanda_res,"Demanda Pedido","Saldo Real","Status"] if c in df_res.columns]
    st.markdown("**Visão geral** — clique nos cards para filtrar por status")
    st.dataframe(df_res[col_ex].style.apply(estilo_linha, axis=1), use_container_width=True, height=380)

# ============================================================
# FILTROS
# ============================================================
elif bloco == "filtros":
    st.markdown('<div class="secao-titulo">🔍 Filtros e Análise Geral</div>', unsafe_allow_html=True)
    st.markdown('<div class="bloco-card">', unsafe_allow_html=True)
    st.markdown('<div class="bloco-title">⚙️ O que considerar na análise</div>', unsafe_allow_html=True)
    col_cfg1, col_cfg2, col_cfg3 = st.columns(3)
    with col_cfg1:
        idx_d = opcoes_almox.index("Saldo Almox 3") if "Saldo Almox 3" in opcoes_almox else 0
        almox_demanda = st.selectbox("Almox — Demanda/Pedidos", opcoes_almox, index=idx_d, key="flt_almox_dem")
    with col_cfg2:
        idx_s = opcoes_almox.index("Saldo Almox 30") if "Saldo Almox 30" in opcoes_almox else min(1, len(opcoes_almox)-1)
        almox_seg = st.selectbox("Almox — Estoque de Segurança", opcoes_almox, index=idx_s, key="flt_almox_seg")
    with col_cfg3:
        usar_ordens = st.checkbox("📋 Incluir Ordens em Aberto", value=True, key="flt_ordens") if tem_ordens else False
    st.markdown('</div>', unsafe_allow_html=True)

    saldo_real = df_full[almox_demanda].copy()
    if usar_ordens and tem_ordens: saldo_real = saldo_real + df_full["Qtde Ordens Abertas"]
    df_calc = df_full.copy()
    df_calc["Saldo vs Demanda"] = df_calc[almox_demanda] - df_calc["Demanda Pedido"]
    df_calc["Saldo Real"] = saldo_real - df_calc["Demanda Pedido"] - df_calc["Qtde Pendente OP"]
    df_calc["Abaixo Estq Seg"] = df_calc[almox_seg] < df_calc["Estq Seg"]

    def calc_status_flt(row):
        if row["Saldo Real"] < 0: return "FALTA"
        if row["Abaixo Estq Seg"]: return "RISCO"
        if row["Demanda Pedido"] + row["Qtde Pendente OP"] >= row[almox_seg] * 0.5: return "RISCO"
        return "OK"

    df_calc["Status"] = df_calc.apply(calc_status_flt, axis=1)

    st.markdown('<div class="bloco-card">', unsafe_allow_html=True)
    st.markdown('<div class="bloco-title">🔎 Pesquisa e Filtros</div>', unsafe_allow_html=True)
    col_f1, col_f2 = st.columns([2,3])
    with col_f1: busca = st.text_input("Pesquisar código ou descrição", placeholder="Ex: ABC123...")
    with col_f2:
        todos_codigos = sorted(df_calc["Codigo"].astype(str).unique().tolist())
        itens_excluir = st.multiselect("🚫 Ocultar itens", options=todos_codigos, placeholder="Selecione para ocultar...")
    col_s1, col_s2, col_s3, col_s4 = st.columns(4)
    for col, key, label in [(col_s1,"TODOS","Todos"),(col_s2,"FALTA","🔴 Falta"),(col_s3,"RISCO","🟡 Risco"),(col_s4,"OK","🟢 OK")]:
        with col:
            if st.button(label, key=f"flt_btn_{key}", use_container_width=True): st.session_state.filtro_status = key
    st.markdown('</div>', unsafe_allow_html=True)

    df_filtrado = df_calc.copy()
    if itens_excluir: df_filtrado = df_filtrado[~df_filtrado["Codigo"].astype(str).isin(itens_excluir)]
    if st.session_state.filtro_status != "TODOS": df_filtrado = df_filtrado[df_filtrado["Status"] == st.session_state.filtro_status]
    if busca: df_filtrado = df_filtrado[df_filtrado["Codigo"].astype(str).str.contains(busca, case=False, na=False) | df_filtrado["Descricao"].astype(str).str.contains(busca, case=False, na=False)]

    colunas_exibir = ["Codigo","Descricao",almox_demanda]
    if almox_seg != almox_demanda: colunas_exibir.append(almox_seg)
    colunas_exibir += ["Demanda Pedido","Qtde Pendente OP"]
    if tem_ordens: colunas_exibir.append("Qtde Ordens Abertas")
    colunas_exibir += ["Saldo vs Demanda","Saldo Real"]
    if tem_parametros: colunas_exibir.append("Estq Seg")
    colunas_exibir.append("Status")
    colunas_exibir = [c for c in colunas_exibir if c in df_filtrado.columns]

    pills = [f"Demanda: {almox_demanda}", f"Est.Seg: {almox_seg}"]
    if usar_ordens and tem_ordens: pills.append("+ Ordens em Aberto")
    if itens_excluir: pills.append(f"🚫 {len(itens_excluir)} oculto(s)")
    st.markdown(" ".join([f'<span class="info-pill">{p}</span>' for p in pills]), unsafe_allow_html=True)
    st.markdown(f"**{len(df_filtrado)} item(ns) exibido(s)**")
    st.dataframe(df_filtrado[colunas_exibir].style.apply(estilo_linha, axis=1), use_container_width=True, height=420)

    col_d1, col_d2 = st.columns(2)
    with col_d1: st.download_button("📥 Baixar CSV", df_filtrado[colunas_exibir].to_csv(index=False).encode("utf-8"), "pcp_filtrado.csv", mime="text/csv", key="dl_flt_csv")
    with col_d2: st.download_button("📥 Baixar Excel", exportar_excel_formatado(df_filtrado[colunas_exibir]), "pcp_filtrado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_flt_xlsx")

# ============================================================
# SALDOS
# ============================================================
elif bloco == "saldos":
    st.markdown('<div class="secao-titulo">💰 Saldos por Almoxarifado</div>', unsafe_allow_html=True)
    st.markdown('<div class="bloco-card">', unsafe_allow_html=True)
    st.markdown('<div class="bloco-title">⚙️ Configurar visualização</div>', unsafe_allow_html=True)
    col_s1, col_s2 = st.columns(2)
    with col_s1: almoxs_exibir = st.multiselect("Quais almoxarifados exibir", options=opcoes_almox, default=opcoes_almox)
    with col_s2: almox_ref = st.selectbox("Almoxarifado de referência", options=opcoes_almox, index=0, key="saldo_ref")
    col_s3, col_s4 = st.columns(2)
    with col_s3: mostrar_zero = st.checkbox("Mostrar itens com saldo zero", value=False)
    with col_s4: busca_saldo = st.text_input("🔍 Pesquisar item", placeholder="Código ou descrição...", key="busca_saldo")
    st.markdown('</div>', unsafe_allow_html=True)

    df_saldo = base.merge(saldo_base, on="Codigo", how="left").fillna(0)
    if not mostrar_zero: df_saldo = df_saldo[df_saldo[almox_ref] > 0]
    if busca_saldo: df_saldo = df_saldo[df_saldo["Codigo"].astype(str).str.contains(busca_saldo, case=False, na=False) | df_saldo["Descricao"].astype(str).str.contains(busca_saldo, case=False, na=False)]
    cols_exibir = ["Codigo","Descricao"] + [a for a in almoxs_exibir if a in df_saldo.columns]

    mc1, mc2, mc3 = st.columns(3)
    mc1.metric("Total com saldo", len(df_saldo[df_saldo[almox_ref] > 0]))
    mc2.metric(f"Soma {almox_ref}", f"{df_saldo[almox_ref].sum():,.0f}")
    mc3.metric("Itens sem saldo", len(df_saldo[df_saldo[almox_ref] == 0]))

    st.markdown(f"**{len(df_saldo)} item(ns)**")
    st.dataframe(df_saldo[cols_exibir], use_container_width=True, height=420)
    col_d1, col_d2 = st.columns(2)
    with col_d1: st.download_button("📥 Baixar Saldos CSV", df_saldo[cols_exibir].to_csv(index=False).encode("utf-8"), "saldos.csv", mime="text/csv", key="dl_saldo_csv")
    with col_d2: st.download_button("📥 Baixar Saldos Excel", exportar_excel_formatado(df_saldo[cols_exibir]), "saldos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_saldo_xlsx")

# ============================================================
# ESTOQUES
# ============================================================
elif bloco == "estoques":
    st.markdown('<div class="secao-titulo">📦 Estoque de Segurança</div>', unsafe_allow_html=True)
    if not tem_parametros:
        st.info("ℹ️ Importe o arquivo de Parâmetros para acessar análise de Estoque de Segurança.")
        st.stop()

    st.markdown('<div class="bloco-card">', unsafe_allow_html=True)
    st.markdown('<div class="bloco-title">⚙️ Configurar análise</div>', unsafe_allow_html=True)
    col_e1, col_e2, col_e3 = st.columns(3)
    with col_e1: almox_estq = st.selectbox("Almox para comparar com Estq. Seg.", options=opcoes_almox, index=min(1,len(opcoes_almox)-1), key="estq_almox")
    with col_e2: incluir_pendente_op = st.checkbox("Somar Qtde Pendente OP ao saldo", value=False, key="estq_inc_op")
    with col_e3: incluir_ordens_estq = st.checkbox("Somar Ordens em Aberto ao saldo", value=False, key="estq_inc_ordens") if tem_ordens else False
    col_e4, col_e5 = st.columns(2)
    with col_e4: apenas_abaixo = st.checkbox("Mostrar apenas itens ABAIXO do estoque de segurança", value=False)
    with col_e5: busca_estq = st.text_input("🔍 Pesquisar", placeholder="Código ou descrição...", key="busca_estq")
    ordem_estq = st.selectbox("Ordenar por", ["Diferença (maior déficit primeiro)","Código","Status"], key="ord_estq")
    st.markdown('</div>', unsafe_allow_html=True)

    df_estq = df_full[["Codigo","Descricao",almox_estq,"Demanda Pedido","Qtde Pendente OP","Estq Seg"]].copy()
    if tem_ordens: df_estq["Qtde Ordens Abertas"] = df_full["Qtde Ordens Abertas"]
    saldo_efetivo = df_estq[almox_estq].copy()
    if incluir_pendente_op: saldo_efetivo = saldo_efetivo + df_estq["Qtde Pendente OP"]
    if incluir_ordens_estq and tem_ordens: saldo_efetivo = saldo_efetivo + df_estq["Qtde Ordens Abertas"]
    df_estq["Saldo Efetivo"] = saldo_efetivo
    df_estq["Diferença"] = df_estq["Saldo Efetivo"] - df_estq["Estq Seg"]
    df_estq["Status Estq"] = df_estq["Diferença"].apply(lambda x: "ABAIXO" if x < 0 else "OK")
    if apenas_abaixo: df_estq = df_estq[df_estq["Status Estq"] == "ABAIXO"]
    if busca_estq: df_estq = df_estq[df_estq["Codigo"].astype(str).str.contains(busca_estq, case=False, na=False) | df_estq["Descricao"].astype(str).str.contains(busca_estq, case=False, na=False)]
    if ordem_estq == "Diferença (maior déficit primeiro)": df_estq = df_estq.sort_values("Diferença")
    elif ordem_estq == "Código": df_estq = df_estq.sort_values("Codigo")
    elif ordem_estq == "Status": df_estq = df_estq.sort_values("Status Estq")

    me1, me2, me3 = st.columns(3)
    me1.metric("⬇️ Abaixo do Estq. Seg.", (df_estq["Status Estq"]=="ABAIXO").sum())
    me2.metric("✅ Acima do Estq. Seg.", (df_estq["Status Estq"]=="OK").sum())
    me3.metric("📊 Almox analisado", almox_estq)

    cols_estq = [c for c in ["Codigo","Descricao",almox_estq,"Saldo Efetivo","Estq Seg","Diferença","Status Estq"] if c in df_estq.columns]

    def estilo_estq(row):
        if row["Status Estq"] == "ABAIXO": return ["background-color:#fee2e2; color:#7f1d1d"] * len(row)
        return ["background-color:#f0fdf4; color:#14532d"] * len(row)

    pills = [f"Almox: {almox_estq}"]
    if incluir_pendente_op: pills.append("+ Pendente OP")
    if incluir_ordens_estq and tem_ordens: pills.append("+ Ordens em Aberto")
    st.markdown(" ".join([f'<span class="info-pill">{p}</span>' for p in pills]), unsafe_allow_html=True)
    st.markdown(f"**{len(df_estq)} item(ns) exibido(s)**")
    st.dataframe(df_estq[cols_estq].style.apply(estilo_estq, axis=1), use_container_width=True, height=420)
    col_d1, col_d2 = st.columns(2)
    with col_d1: st.download_button("📥 Baixar CSV", df_estq[cols_estq].to_csv(index=False).encode("utf-8"), "estoques.csv", mime="text/csv", key="dl_estq_csv")
    with col_d2: st.download_button("📥 Baixar Excel", exportar_excel_formatado(df_estq[cols_estq]), "estoques.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_estq_xlsx")

# ============================================================
# DEMANDAS
# ============================================================
elif bloco == "demandas":
    st.markdown('<div class="secao-titulo">📋 Análise de Demandas</div>', unsafe_allow_html=True)
    st.markdown('<div class="bloco-card">', unsafe_allow_html=True)
    st.markdown('<div class="bloco-title">⚙️ O que incluir na análise de demanda</div>', unsafe_allow_html=True)
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        almox_dem = st.selectbox("Almox de referência para saldo", options=opcoes_almox, index=0, key="dem_almox")
        incluir_dc = st.checkbox("Incluir Demanda de Pedidos (DC)", value=True, key="dem_dc")
        incluir_op_dem = st.checkbox("Incluir Qtde Pendente OP", value=True, key="dem_op")
    with col_d2:
        incluir_ord_dem = st.checkbox("Incluir Ordens em Aberto", value=True, key="dem_ordens") if tem_ordens else False
        apenas_com_demanda = st.checkbox("Mostrar apenas itens com demanda > 0", value=False, key="dem_apenas")
        busca_dem = st.text_input("🔍 Pesquisar", placeholder="Código ou descrição...", key="busca_dem")
    st.markdown('</div>', unsafe_allow_html=True)

    df_dem = df_full[["Codigo","Descricao",almox_dem,"Demanda Pedido","Qtde Pendente OP"]].copy()
    if tem_ordens: df_dem["Qtde Ordens Abertas"] = df_full["Qtde Ordens Abertas"]
    demanda_total = pd.Series([0.0]*len(df_dem), index=df_dem.index)
    if incluir_dc: demanda_total = demanda_total + df_dem["Demanda Pedido"]
    if incluir_op_dem: demanda_total = demanda_total + df_dem["Qtde Pendente OP"]
    if incluir_ord_dem and tem_ordens: demanda_total = demanda_total + df_dem["Qtde Ordens Abertas"]
    df_dem["Demanda Total"] = demanda_total
    df_dem["Saldo vs Demanda"] = df_dem[almox_dem] - df_dem["Demanda Total"]
    df_dem["Cobertura %"] = (df_dem[almox_dem] / df_dem["Demanda Total"].replace(0, float("nan")) * 100).fillna(0).round(1)
    df_dem["Status"] = df_dem["Saldo vs Demanda"].apply(lambda x: "FALTA" if x < 0 else "OK")
    if apenas_com_demanda: df_dem = df_dem[df_dem["Demanda Total"] > 0]
    if busca_dem: df_dem = df_dem[df_dem["Codigo"].astype(str).str.contains(busca_dem, case=False, na=False) | df_dem["Descricao"].astype(str).str.contains(busca_dem, case=False, na=False)]
    df_dem = df_dem.sort_values("Saldo vs Demanda")

    md1, md2, md3, md4 = st.columns(4)
    md1.metric("📦 Total itens", len(df_dem))
    md2.metric("🔴 Com falta", (df_dem["Status"]=="FALTA").sum())
    md3.metric("📋 Demanda total", f"{df_dem['Demanda Total'].sum():,.0f}")
    md4.metric("💰 Saldo total", f"{df_dem[almox_dem].sum():,.0f}")

    cols_dem = ["Codigo","Descricao",almox_dem,"Demanda Pedido"]
    if incluir_op_dem: cols_dem.append("Qtde Pendente OP")
    if incluir_ord_dem and tem_ordens: cols_dem.append("Qtde Ordens Abertas")
    cols_dem += ["Demanda Total","Saldo vs Demanda","Cobertura %","Status"]
    cols_dem = [c for c in cols_dem if c in df_dem.columns]

    def estilo_dem(row):
        if row["Status"] == "FALTA": return ["background-color:#fee2e2; color:#7f1d1d"] * len(row)
        return ["background-color:#f0fdf4; color:#14532d"] * len(row)

    pills_dem = [f"Almox: {almox_dem}"]
    if incluir_dc: pills_dem.append("DC")
    if incluir_op_dem: pills_dem.append("Pendente OP")
    if incluir_ord_dem and tem_ordens: pills_dem.append("Ordens em Aberto")
    st.markdown(" ".join([f'<span class="info-pill">{p}</span>' for p in pills_dem]), unsafe_allow_html=True)
    st.markdown(f"**{len(df_dem)} item(ns)**")
    st.dataframe(df_dem[cols_dem].style.apply(estilo_dem, axis=1), use_container_width=True, height=420)
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1: st.download_button("📥 Baixar CSV", df_dem[cols_dem].to_csv(index=False).encode("utf-8"), "demandas.csv", mime="text/csv", key="dl_dem_csv")
    with col_dl2: st.download_button("📥 Baixar Excel", exportar_excel_formatado(df_dem[cols_dem]), "demandas.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_dem_xlsx")

# ============================================================
# SUGESTÃO
# ============================================================
elif bloco == "sugestao":
    st.markdown('<div class="secao-titulo">🏭 Sugestão de Produção</div>', unsafe_allow_html=True)
    st.markdown('<div class="bloco-card">', unsafe_allow_html=True)
    st.markdown('<div class="bloco-title">⚙️ Configurar base para sugestão</div>', unsafe_allow_html=True)
    col_sg1, col_sg2, col_sg3 = st.columns(3)
    with col_sg1:
        idx_sug = opcoes_almox.index("Saldo Almox 3") if "Saldo Almox 3" in opcoes_almox else 0
        almox_sug = st.selectbox("Almox de referência", opcoes_almox, index=idx_sug, key="sug_almox")
    with col_sg2: incluir_ordens_sug = st.checkbox("Incluir Ordens em Aberto no saldo", value=True, key="sug_ordens") if tem_ordens else False
    with col_sg3: incluir_estq_seg_sug = st.checkbox("Incluir Estq. Seg. na sugestão", value=True, key="sug_estq") if tem_parametros else False
    st.markdown('</div>', unsafe_allow_html=True)

    saldo_sug = df_full[almox_sug].copy()
    if incluir_ordens_sug and tem_ordens: saldo_sug = saldo_sug + df_full["Qtde Ordens Abertas"]
    df_sug = df_full.copy()
    df_sug["Saldo Real"] = saldo_sug - df_sug["Demanda Pedido"] - df_sug["Qtde Pendente OP"]
    df_sug["Abaixo Estq Seg"] = df_sug[almox_sug] < df_sug["Estq Seg"]

    def calc_status_sug(row):
        if row["Saldo Real"] < 0: return "FALTA"
        if row["Abaixo Estq Seg"]: return "RISCO"
        if row["Demanda Pedido"] + row["Qtde Pendente OP"] >= row[almox_sug] * 0.5: return "RISCO"
        return "OK"

    df_sug["Status"] = df_sug.apply(calc_status_sug, axis=1)
    df_sug_filt = df_sug[df_sug["Status"].isin(["FALTA","RISCO"])].copy()

    if not df_sug_filt.empty:
        df_sug_filt["Qtde Sugerida"] = df_sug_filt.apply(
            lambda row: max(row["Demanda Pedido"] + row["Qtde Pendente OP"] - row[almox_sug] + (row["Estq Seg"] if incluir_estq_seg_sug and tem_parametros else 0), 0), axis=1)
        df_sug_filt = df_sug_filt[df_sug_filt["Qtde Sugerida"] > 0].sort_values(by=["Status","Qtde Sugerida"], ascending=[True,False])
        n_urgente = (df_sug_filt["Status"]=="FALTA").sum(); n_atencao = (df_sug_filt["Status"]=="RISCO").sum()

        col_sg_a, col_sg_b, col_sg_c = st.columns([3,1,1])
        with col_sg_a:
            st.markdown(f'<div class="sugestao-box"><div style="font-size:15px;font-weight:700;margin-bottom:14px;">📋 {len(df_sug_filt)} ITEM(NS) PRECISAM SER PRODUZIDOS</div>', unsafe_allow_html=True)
            for _, row in df_sug_filt.head(10).iterrows():
                cls = "urgente" if row["Status"]=="FALTA" else "atencao"
                emoji = "🔴" if row["Status"]=="FALTA" else "🟡"
                st.markdown(f'<div class="sugestao-item {cls}"><div><div style="font-weight:600;font-size:13px;">{emoji} {row["Codigo"]}</div><div style="font-size:12px;opacity:.75;">{str(row["Descricao"])[:50]}</div></div><div style="background:rgba(255,255,255,.2);border-radius:8px;padding:4px 10px;font-weight:700;font-size:13px;">Prod: {row["Qtde Sugerida"]:,.0f}</div></div>', unsafe_allow_html=True)
            if len(df_sug_filt) > 10: st.markdown(f'<div style="color:rgba(255,255,255,.6);font-size:12px;margin-top:8px;">+ {len(df_sug_filt)-10} outros (veja o Excel)</div>', unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)
        with col_sg_b: st.metric("🔴 Urgente", n_urgente)
        with col_sg_c: st.metric("🟡 Atenção", n_atencao)

        with st.expander("📄 Ver tabela completa", expanded=False):
            cols_sug_t = [c for c in ["Codigo","Descricao",almox_sug,"Demanda Pedido","Qtde Pendente OP","Estq Seg","Qtde Sugerida","Status"] if c in df_sug_filt.columns]
            st.dataframe(df_sug_filt[cols_sug_t].style.apply(estilo_linha, axis=1), use_container_width=True)

        df_down_sug = df_sug_filt[[c for c in ["Codigo","Descricao",almox_sug,"Demanda Pedido","Qtde Pendente OP","Estq Seg","Qtde Sugerida","Status"] if c in df_sug_filt.columns]].copy()
        st.download_button("📥 Baixar Sugestão (Excel)", exportar_excel_formatado(df_down_sug), "sugestao_producao.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_sugestao")
    else:
        st.success("✅ Nenhum item necessita produção. Todos os saldos estão adequados.")
