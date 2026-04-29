import re
from datetime import datetime
from io import BytesIO

import pandas as pd
import pytz
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

st.set_page_config(page_title="Monitoramento PCP", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Segoe+UI:wght@300;400;600;700;800&display=swap');
*, *::before, *::after { box-sizing: border-box; }
html, body, [class*="css"] { font-family: 'Segoe UI', Tahoma, sans-serif; font-size: 11px; }
.main { background: #d4d8dc; }
.block-container { padding-top: 0.5rem !important; padding-bottom: 1rem !important; max-width: 100% !important; }

/* HEADER */
.pcp-header { background: #c8ccce; border: 1px solid #999; border-bottom: 2px solid #888; padding: 6px 14px; display: flex; align-items: center; justify-content: space-between; margin-bottom: 6px; }
.pcp-logo { font-size: 22px; font-weight: 900; color: #c0392b; font-style: italic; letter-spacing: -1px; }
.pcp-title { font-size: 15px; font-weight: 700; color: #111; text-align: center; flex: 1; }
.pcp-date-section { text-align: right; min-width: 120px; }
.pcp-date-label { font-size: 9px; color: #555; margin-bottom: 2px; border-bottom: 1px solid #aaa; padding-bottom: 1px; }
.pcp-date-val { font-size: 11px; font-weight: 700; color: #1a1a2e; padding: 2px 0; }

/* PANELS */
.panel { background: #e8eaec; border: 1px solid #aaa; overflow: hidden; margin-bottom: 6px; }
.panel-header { background: #c0c4c6; border-bottom: 1px solid #999; padding: 3px 8px; font-size: 10px; font-weight: 700; color: #111; text-align: center; letter-spacing: .3px; }

/* STATUS CARDS */
.status-cards { display: grid; grid-template-columns: repeat(4,1fr); gap: 5px; padding: 6px; }
.scard { padding: 6px 8px; border-radius: 3px; cursor: pointer; border: 1px solid rgba(0,0,0,.1); }
.scard-total { background: #1a3a5c; color: white; }
.scard-falta { background: #c0392b; color: white; }
.scard-risco { background: #d68910; color: white; }
.scard-ok    { background: #1e8449; color: white; }
.scard-lbl { font-size: 9px; font-weight: 600; opacity: .85; text-transform: uppercase; letter-spacing: .5px; }
.scard-num { font-size: 24px; font-weight: 800; line-height: 1; }
.scard-sub { font-size: 8px; opacity: .75; margin-top: 1px; }

/* BADGES */
.badge { display: inline-flex; align-items: center; gap: 2px; padding: 1px 5px; border-radius: 2px; font-size: 9px; font-weight: 700; }
.badge-ok    { background: #c6efce; color: #1e6b2e; border: 1px solid #9dc09d; }
.badge-risco { background: #ffeb9c; color: #7a5000; border: 1px solid #d4b800; }
.badge-falta { background: #ffc7ce; color: #9c0006; border: 1px solid #e0808a; }

/* TABLES */
div[data-testid="stDataFrame"] { background: #e8eaec !important; border: 1px solid #aaa !important; border-radius: 0 !important; box-shadow: none !important; font-size: 9.5px !important; }

/* FILTER BAR */
.filter-info { display: flex; gap: 6px; padding: 5px 6px; align-items: center; background: #dde0e3; border-bottom: 1px solid #bbb; font-size: 10px; }

/* METRICS */
.metric-row { display: flex; gap: 8px; padding: 5px 6px; background: #dde0e3; border-bottom: 1px solid #ccc; margin-bottom: 4px; }
.metric-box { background: #cdd0d3; border: 1px solid #bbb; border-radius: 2px; padding: 4px 8px; min-width: 80px; }
.metric-lbl { font-size: 9px; color: #555; font-weight: 700; }
.metric-val { font-size: 14px; font-weight: 800; color: #1a1a2e; }

/* BUTTONS */
div.stButton > button { border-radius: 2px !important; font-weight: 600 !important; font-size: 10px !important; border: 1px solid #999 !important; background: #d0d4d8 !important; color: #222 !important; transition: all .1s !important; padding: 4px 10px !important; }
div.stButton > button:hover { background: #bbbfc3 !important; border-color: #777 !important; }
div.stButton > button[kind="primary"] { background: #1a3a5c !important; color: white !important; border-color: #0d2740 !important; }

/* NAV TABS */
div[data-baseweb="tab-list"] { background: #c8ccce; border-bottom: 2px solid #999; gap: 2px; }
div[data-baseweb="tab"] { background: #c8ccce; border: 1px solid #999; border-bottom: none; border-radius: 2px 2px 0 0; font-size: 10px; font-weight: 600; color: #555; padding: 4px 12px; }
div[aria-selected="true"][data-baseweb="tab"] { background: #e8eaec; color: #111; border-bottom: 2px solid #e8eaec; }

/* SELECTS / INPUTS */
div[data-baseweb="select"] > div { border-radius: 2px !important; border: 1px solid #aaa !important; background: #f0f2f4 !important; font-size: 10px !important; }
div.stTextInput > div > input { border-radius: 2px !important; border: 1px solid #aaa !important; background: #f0f2f4 !important; font-size: 10px !important; }

/* INFO PILLS */
.info-pill { display: inline-flex; align-items: center; gap: 6px; background: #dde0e3; border: 1px solid #bbb; border-radius: 2px; padding: 2px 8px; font-size: 10px; color: #333; font-weight: 600; margin-right: 4px; margin-bottom: 4px; }

/* SUGESTAO ITEMS */
.sug-item { display: flex; justify-content: space-between; align-items: center; padding: 3px 8px; border-bottom: 1px solid #ccc; font-size: 9.5px; background: #e8eaec; }
.sug-item-falta { border-left: 3px solid #c0392b; }
.sug-item-risco { border-left: 3px solid #d68910; }
.sug-cod { font-weight: 700; color: #111; }
.sug-qtd { font-weight: 700; color: #1a3a5c; background: rgba(0,0,0,.08); border-radius: 2px; padding: 2px 6px; }

/* MSG */
.msg-warn { background: #fff3cd; border: 1px solid #e6a817; color: #7a5000; padding: 6px 10px; font-size: 10px; border-radius: 2px; margin: 4px 0; }
.msg-ok   { background: #d4edda; border: 1px solid #7dbb8e; color: #155724; padding: 6px 10px; font-size: 10px; border-radius: 2px; margin: 4px 0; }
.msg-info { background: #d0e8f8; border: 1px solid #7ab3d4; color: #0c4a6e; padding: 6px 10px; font-size: 10px; border-radius: 2px; margin: 4px 0; }

/* SCROLLBAR */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: #d0d4d8; }
::-webkit-scrollbar-thumb { background: #999; border-radius: 3px; }
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
            ws.column_dimensions[letra].width = min(
                max(len(str(coluna)) + 4,
                    max((len(str(v)) for v in df[coluna] if pd.notna(v)), default=4) + 2), 40)
        if "Status" in df.columns:
            idx_s = list(df.columns).index("Status") + 1
            for cell in ws[get_column_letter(idx_s)][1:]:
                v = str(cell.value or "")
                if "FALTA" in v:  cell.fill = PatternFill(fill_type="solid", fgColor="FFC7CE")
                elif "RISCO" in v: cell.fill = PatternFill(fill_type="solid", fgColor="FFF2CC")
                elif "OK" in v:   cell.fill = PatternFill(fill_type="solid", fgColor="C6E0B4")
    output.seek(0)
    return output.getvalue()

def badge_html(status):
    if status == "OK":    return '<span class="badge badge-ok">✔ OK</span>'
    if status == "RISCO": return '<span class="badge badge-risco">⚠ RISCO</span>'
    if status == "FALTA": return '<span class="badge badge-falta">● FALTA</span>'
    return status

def metric_box(lbl, val):
    return f'<div class="metric-box"><div class="metric-lbl">{lbl}</div><div class="metric-val">{val}</div></div>'

def fmt_num(n, dec=0):
    try:
        return f"{float(n):,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(n)

def estilo_linha(row):
    if row["Status"] == "FALTA": return ["background-color:#ffc7ce; color:#9c0006"] * len(row)
    if row["Status"] == "RISCO": return ["background-color:#ffeb9c; color:#7a5000"] * len(row)
    if row["Status"] == "OK":   return ["background-color:#f0fff4; color:#14532d"] * len(row)
    return [""] * len(row)

# ── CARREGAR DADOS ──────────────────────────────────────────
try:
    saldo    = pd.read_csv("saldo.csv")
    perfil   = pd.read_csv("perfil.csv")
    previsao = pd.read_csv("previsao.csv")
except Exception:
    st.markdown('<div class="msg-warn">⚠️ Faça upload dos dados na página principal primeiro.</div>', unsafe_allow_html=True)
    st.stop()

tem_ordens = False
ordens_abertas = pd.DataFrame()
try:
    ordens_raw = pd.read_csv("ordens.csv")
    colunas_upper = {c: c.upper().strip() for c in ordens_raw.columns}
    col_cod_ordens = col_qtd_ordens = col_status_ordens = None
    for orig, upper in colunas_upper.items():
        if any(k in upper for k in ["COD","ITEM","PRODUTO","CODIGO","CÓDIGO"]) and col_cod_ordens is None: col_cod_ordens = orig
        if any(k in upper for k in ["QTDE","QTD","QUANTIDADE","SALDO","SALDO ORDEM"]) and col_qtd_ordens is None: col_qtd_ordens = orig
        if any(k in upper for k in ["STATUS","SITUAÇÃO","SITUACAO","SIT"]) and col_status_ordens is None: col_status_ordens = orig
    if col_cod_ordens and col_qtd_ordens:
        cols = [col_cod_ordens, col_qtd_ordens] + ([col_status_ordens] if col_status_ordens else [])
        df_ord = ordens_raw[cols].copy()
        df_ord.columns = ["Codigo","Qtde Ordem"] + (["Status Ordem"] if col_status_ordens else [])
        if "Status Ordem" in df_ord.columns:
            palavras_aberto = ["ABERTO","ABERTA","EM ABERTO","LIBERADA","LIBERADO","PENDENTE","A PRODUZIR"]
            mask = df_ord["Status Ordem"].astype(str).str.upper().str.strip().isin(palavras_aberto)
            if mask.sum() > 0: df_ord = df_ord[mask]
        df_ord["Qtde Ordem"] = pd.to_numeric(
            df_ord["Qtde Ordem"].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
            errors="coerce").fillna(0)
        ordens_abertas = df_ord.groupby("Codigo")["Qtde Ordem"].sum().reset_index()
        ordens_abertas.columns = ["Codigo","Qtde Ordens Abertas"]
        tem_ordens = True
except Exception:
    pass

tem_parametros = False
try:
    parametros = pd.read_csv("parametros.csv")
    parametros = parametros.rename(columns={"COD ITEM": "Codigo", "ESTQ SEG": "Estq Seg"})
    parametros = parametros[["Codigo","Estq Seg"]].copy()
    parametros["Estq Seg"] = pd.to_numeric(parametros["Estq Seg"], errors="coerce").fillna(0)
    tem_parametros = True
except Exception:
    pass

saldo    = saldo.sort_values(by=["Data Processamento","Hora Processamento"], ascending=False).drop_duplicates("Codigo")
previsao = previsao.sort_values(by=["Data Processamento","Hora Processamento"], ascending=False).drop_duplicates("COD")
base     = previsao[["COD","PRODUTO"]].copy()
base.columns = ["Codigo","Descricao"]

colunas_almox = [c for c in saldo.columns if "SALDO" in c.upper() and "ALMOX" in c.upper()]
opcoes_almox  = (["Saldo Total"] if "Saldo Total" in saldo.columns else []) + colunas_almox
if not opcoes_almox:
    st.error("Nenhuma coluna de saldo encontrada. Reprocesse o PDF de Saldo.")
    st.stop()
for col in opcoes_almox:
    if col not in saldo.columns: saldo[col] = 0

saldo_base = saldo[["Codigo"] + opcoes_almox]
perfil["Quantidade"] = pd.to_numeric(
    perfil["Quantidade"].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
    errors="coerce").fillna(0)

dc = perfil[perfil["Tipo"] == "DC"].groupby("Item")["Quantidade"].sum().reset_index()
dc.columns = ["Codigo","Demanda Pedido"]
op = perfil[perfil["Tipo"].isin(["OP","ORDEM","LIBERADA"])].groupby("Item")["Quantidade"].sum().reset_index()
op.columns = ["Codigo","Qtde Pendente OP"]

df_full = base.merge(saldo_base, on="Codigo", how="left")
df_full = df_full.merge(dc, on="Codigo", how="left")
df_full = df_full.merge(op, on="Codigo", how="left")
if tem_ordens:     df_full = df_full.merge(ordens_abertas, on="Codigo", how="left")
if tem_parametros:
    df_full = df_full.merge(parametros, on="Codigo", how="left")
    df_full["Estq Seg"] = df_full["Estq Seg"].fillna(0)
else:
    df_full["Estq Seg"] = 0
df_full = df_full.fillna(0)

# ── HEADER ──────────────────────────────────────────────────
now_str = agora().strftime("%d/%m/%Y %H:%M")
st.markdown(f"""
<div class="pcp-header">
  <div class="pcp-logo">Colafix</div>
  <div class="pcp-title">Monitoramento PCP</div>
  <div class="pcp-date-section">
    <div class="pcp-date-label">Data Relatório ▼</div>
    <div class="pcp-date-val">{now_str}</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── NAVEGAÇÃO ────────────────────────────────────────────────
tab_upload, tab_dash, tab_filtros, tab_estoques, tab_demandas, tab_sugestao = st.tabs([
    "⬆ Upload de Dados", "📊 Dashboard PCP", "🔍 Filtros",
    "📦 Estoques", "📋 Demandas", "🏭 Sugestão"
])

def calc_status(row, almox_dem, almox_seg, usar_ordens):
    saldo_real = (row[almox_dem] + (row["Qtde Ordens Abertas"] if usar_ordens and tem_ordens else 0)
                  - row["Demanda Pedido"] - row["Qtde Pendente OP"])
    abaixo_seg = row[almox_seg] < row["Estq Seg"]
    if saldo_real < 0: return "FALTA"
    if abaixo_seg:     return "RISCO"
    if row["Demanda Pedido"] + row["Qtde Pendente OP"] >= row[almox_seg] * 0.5: return "RISCO"
    return "OK"

# ══════════════════════════════════════════════════════════════
# ABA UPLOAD
# ══════════════════════════════════════════════════════════════
with tab_upload:
    st.markdown('<div class="panel"><div class="panel-header">Upload e Processamento de Dados</div>', unsafe_allow_html=True)
    st.markdown('<div class="msg-info">ℹ️ Os dados são carregados automaticamente a partir dos arquivos CSV/Excel salvos pelo sistema principal (app.py). Esta aba é apenas informativa.</div>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("**📄 Saldo**")
        st.success(f"✅ {len(saldo)} itens carregados" if len(saldo) > 0 else "❌ Não carregado")
    with col2:
        st.markdown("**📋 Perfil**")
        st.success(f"✅ {len(perfil)} registros carregados" if len(perfil) > 0 else "❌ Não carregado")
    with col3:
        st.markdown("**📈 Previsão**")
        st.success(f"✅ {len(previsao)} itens carregados" if len(previsao) > 0 else "❌ Não carregado")
    col4, col5, col6 = st.columns(3)
    with col4:
        st.markdown("**🏭 Ordens**")
        if tem_ordens: st.success(f"✅ {len(ordens_abertas)} ordens abertas")
        else: st.warning("⚠️ Não carregado")
    with col5:
        st.markdown("**⚙️ Parâmetros**")
        if tem_parametros: st.success(f"✅ {len(parametros)} itens carregados")
        else: st.warning("⚠️ Não carregado")
    with col6:
        st.markdown("**📦 Base consolidada**")
        st.info(f"🔢 {len(df_full)} itens no total")
    st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# ABA DASHBOARD
# ══════════════════════════════════════════════════════════════
with tab_dash:
    almox_d_dash = opcoes_almox[0] if not any("3" in x for x in opcoes_almox) else next(x for x in opcoes_almox if "3" in x)
    almox_s_dash = opcoes_almox[-1] if not any("30" in x for x in opcoes_almox) else next(x for x in opcoes_almox if "30" in x)

    rows_dash = df_full.copy()
    rows_dash["Status"] = rows_dash.apply(lambda r: calc_status(r, almox_d_dash, almox_s_dash, True), axis=1)

    n_total = len(rows_dash)
    n_falta = (rows_dash["Status"] == "FALTA").sum()
    n_risco = (rows_dash["Status"] == "RISCO").sum()
    n_ok    = (rows_dash["Status"] == "OK").sum()

    # Cards resumo
    st.markdown('<div class="panel"><div class="panel-header">Resumo de Status</div>', unsafe_allow_html=True)
    st.markdown(f"""
    <div class="status-cards">
      <div class="scard scard-total"><div class="scard-lbl">📦 Total</div><div class="scard-num">{n_total}</div><div class="scard-sub">itens na base</div></div>
      <div class="scard scard-falta"><div class="scard-lbl">🔴 Falta</div><div class="scard-num">{n_falta}</div><div class="scard-sub">produção urgente</div></div>
      <div class="scard scard-risco"><div class="scard-lbl">🟡 Risco</div><div class="scard-num">{n_risco}</div><div class="scard-sub">abaixo est. seg.</div></div>
      <div class="scard scard-ok">  <div class="scard-lbl">🟢 OK</div>   <div class="scard-num">{n_ok}</div>   <div class="scard-sub">dentro do esperado</div></div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Linha principal: Consumo | Saldo x Pedidos | Pedidos/Sugestão
    col_main, col_right1, col_right2 = st.columns([2.4, 0.75, 0.75])

    with col_main:
        st.markdown('<div class="panel"><div class="panel-header">Consumo</div>', unsafe_allow_html=True)
        busca_c = st.text_input("Buscar código ou descrição...", key="busca_consumo_dash", placeholder="Pesquisar...")
        filtro_st_c = st.selectbox("Filtrar status", ["Todos","FALTA","RISCO","OK"], key="filt_st_consumo_dash")

        df_consumo = rows_dash.copy()
        if busca_c:
            df_consumo = df_consumo[
                df_consumo["Codigo"].astype(str).str.contains(busca_c, case=False, na=False) |
                df_consumo["Descricao"].astype(str).str.contains(busca_c, case=False, na=False)]
        if filtro_st_c != "Todos":
            df_consumo = df_consumo[df_consumo["Status"] == filtro_st_c]

        # Montar tabela HTML compacta
        rows_html = ""
        for _, r in df_consumo.head(80).iterrows():
            cls = "row-falta" if r["Status"]=="FALTA" else ("row-risco" if r["Status"]=="RISCO" else "")
            bg  = "background:#ffc7ce" if r["Status"]=="FALTA" else ("background:#ffeb9c" if r["Status"]=="RISCO" else "background:#f0fff4")
            badge = badge_html(r["Status"])
            rows_html += f"""<tr style="{bg}">
              <td><b>{r['Codigo']}</b></td>
              <td style="text-align:right">{fmt_num(r.get('Saldo Total',0))}</td>
              <td style="max-width:160px;white-space:normal;font-size:9px">{str(r['Descricao'])[:50]}</td>
              <td style="text-align:right">{fmt_num(r.get(almox_d_dash,0))}</td>
              <td style="text-align:right">{fmt_num(r['Demanda Pedido'])}</td>
              <td>{badge}</td>
            </tr>"""
        st.markdown(f"""
        <div style="overflow:auto;max-height:280px;border:1px solid #ccc">
        <table style="width:100%;border-collapse:collapse;font-size:9.5px">
          <thead><tr>
            <th style="background:#c0c4c6;border:1px solid #999;padding:3px 5px;font-weight:700;position:sticky;top:0">Codigo</th>
            <th style="background:#c0c4c6;border:1px solid #999;padding:3px 5px;font-weight:700;position:sticky;top:0">Saldo Total</th>
            <th style="background:#c0c4c6;border:1px solid #999;padding:3px 5px;font-weight:700;position:sticky;top:0">Descrição</th>
            <th style="background:#c0c4c6;border:1px solid #999;padding:3px 5px;font-weight:700;position:sticky;top:0">{almox_d_dash}</th>
            <th style="background:#c0c4c6;border:1px solid #999;padding:3px 5px;font-weight:700;position:sticky;top:0">Demanda Pedido</th>
            <th style="background:#c0c4c6;border:1px solid #999;padding:3px 5px;font-weight:700;position:sticky;top:0">Status</th>
          </tr></thead>
          <tbody>{rows_html or '<tr><td colspan="6" style="text-align:center;color:#888;padding:12px">Sem dados</td></tr>'}</tbody>
        </table></div>
        """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with col_right1:
        # Saldo x Pedidos (mini chart via barras CSS)
        st.markdown('<div class="panel"><div class="panel-header">Saldo x Pedidos</div>', unsafe_allow_html=True)
        sample = rows_dash.head(12)
        max_v  = max(sample[["Saldo Total","Demanda Pedido"]].max().max(), 1)
        bars_html = ""
        for _, r in sample.iterrows():
            sh = int((r.get("Saldo Total",0)/max_v)*60)
            dh = int((r["Demanda Pedido"]/max_v)*60)
            col_bar = "#c0392b" if r["Status"]=="FALTA" else ("#e6a817" if r["Status"]=="RISCO" else "#27ae60")
            cod_lbl = str(r["Codigo"])[:5]
            bars_html += f"""<div style="display:flex;flex-direction:column;align-items:center;flex:1">
              <div style="display:flex;align-items:flex-end;gap:1px;height:60px">
                <div style="height:{sh}px;background:{col_bar};width:6px;border-radius:1px 1px 0 0"></div>
                <div style="height:{dh}px;background:#666;width:4px;opacity:.5;border-radius:1px 1px 0 0"></div>
              </div>
              <div style="font-size:7px;color:#666;transform:rotate(-40deg);transform-origin:top center;width:18px;overflow:hidden;margin-top:2px">{cod_lbl}</div>
            </div>"""
        st.markdown(f"""
        <div style="font-size:9px;padding:3px 6px">
          <span style="color:#c0392b">●</span> FALTA &nbsp;
          <span style="color:#27ae60">●</span> OK &nbsp;
          <span style="color:#e6a817">●</span> RISCO
        </div>
        <div style="display:flex;align-items:flex-end;gap:2px;height:76px;padding:4px 6px 0;border-bottom:1px solid #bbb">
          {bars_html}
        </div>
        """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # Previsão mês
        st.markdown('<div class="panel"><div class="panel-header">Previsão mês</div>', unsafe_allow_html=True)
        prev_rows = "".join(f'<tr><td><b>{r["COD"]}</b></td><td style="text-align:right;font-size:9px">{str(r["PRODUTO"])[:20]}</td></tr>'
                            for _, r in previsao.head(8).iterrows()) or '<tr><td colspan="2" style="color:#999;text-align:center">—</td></tr>'
        st.markdown(f"""
        <div style="overflow:auto;max-height:100px">
        <table style="width:100%;border-collapse:collapse;font-size:9.5px">
          <thead><tr>
            <th style="background:#c0c4c6;border:1px solid #999;padding:3px 5px;position:sticky;top:0">Codigo</th>
            <th style="background:#c0c4c6;border:1px solid #999;padding:3px 5px;position:sticky;top:0">Produto</th>
          </tr></thead>
          <tbody>{prev_rows}</tbody>
        </table></div>
        """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with col_right2:
        # Pedidos por data
        st.markdown('<div class="panel"><div class="panel-header">Pedidos por data</div>', unsafe_allow_html=True)
        pedidos_dc = perfil[perfil["Tipo"]=="DC"].sort_values("Data Fim").head(10)
        ped_rows = "".join(f'<tr><td><b>{r["Item"]}</b></td><td>{r["Data Fim"]}</td><td style="text-align:right">{fmt_num(r["Quantidade"])}</td></tr>'
                           for _, r in pedidos_dc.iterrows()) or '<tr><td colspan="3" style="color:#999;text-align:center">—</td></tr>'
        st.markdown(f"""
        <div style="overflow:auto;max-height:130px">
        <table style="width:100%;border-collapse:collapse;font-size:9.5px">
          <thead><tr>
            <th style="background:#c0c4c6;border:1px solid #999;padding:3px 5px;position:sticky;top:0">Item</th>
            <th style="background:#c0c4c6;border:1px solid #999;padding:3px 5px;position:sticky;top:0">Data Fim</th>
            <th style="background:#c0c4c6;border:1px solid #999;padding:3px 5px;position:sticky;top:0">Demanda</th>
          </tr></thead>
          <tbody>{ped_rows}</tbody>
        </table></div>
        """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # Sugestão mini
        st.markdown('<div class="panel"><div class="panel-header">Sugestão de Produção</div>', unsafe_allow_html=True)
        df_sug_mini = rows_dash[rows_dash["Status"].isin(["FALTA","RISCO"])].copy()
        df_sug_mini["QtdSug"] = df_sug_mini.apply(
            lambda r: max(r["Demanda Pedido"] + r["Qtde Pendente OP"] - r.get(almox_d_dash,0) + r["Estq Seg"], 0), axis=1)
        df_sug_mini = df_sug_mini[df_sug_mini["QtdSug"] > 0].sort_values("QtdSug", ascending=False)
        sug_rows = "".join(
            f'<tr><td><b>{r["Codigo"]}</b></td><td style="text-align:right">{fmt_num(r["QtdSug"])}</td>'
            f'<td style="color:#c0392b;font-weight:700;font-size:9px">Produzir</td></tr>'
            for _, r in df_sug_mini.head(12).iterrows()
        ) or '<tr><td colspan="3" style="color:#27ae60;text-align:center">Nenhuma sugestão</td></tr>'
        st.markdown(f"""
        <div style="overflow:auto;max-height:140px">
        <table style="width:100%;border-collapse:collapse;font-size:9.5px">
          <thead><tr>
            <th style="background:#c0c4c6;border:1px solid #999;padding:3px 5px;position:sticky;top:0">Codigo</th>
            <th style="background:#c0c4c6;border:1px solid #999;padding:3px 5px;position:sticky;top:0">Sugestão</th>
            <th style="background:#c0c4c6;border:1px solid #999;padding:3px 5px;position:sticky;top:0">Status</th>
          </tr></thead>
          <tbody>{sug_rows}</tbody>
        </table></div>
        """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# ABA FILTROS
# ══════════════════════════════════════════════════════════════
with tab_filtros:
    st.markdown('<div class="panel"><div class="panel-header">Filtros e Análise Geral</div>', unsafe_allow_html=True)

    col_f1, col_f2, col_f3, col_f4, col_f5 = st.columns([1.2, 1.2, 0.8, 1.5, 0.8])
    with col_f1:
        idx_d = opcoes_almox.index("Saldo Almox 3") if "Saldo Almox 3" in opcoes_almox else 0
        almox_dem = st.selectbox("Almox Demanda", opcoes_almox, index=idx_d, key="flt_almox_dem")
    with col_f2:
        idx_s = opcoes_almox.index("Saldo Almox 30") if "Saldo Almox 30" in opcoes_almox else min(1, len(opcoes_almox)-1)
        almox_seg = st.selectbox("Almox Est.Seg", opcoes_almox, index=idx_s, key="flt_almox_seg")
    with col_f3:
        usar_ordens_flt = st.checkbox("Incl. Ordens Abertas", value=True, key="flt_ordens") if tem_ordens else False
    with col_f4:
        busca_flt = st.text_input("Pesquisar código/descrição...", key="busca_flt", placeholder="Pesquisar...")
    with col_f5:
        filtro_st_flt = st.selectbox("Status", ["Todos","FALTA","RISCO","OK"], key="flt_status")

    df_calc = df_full.copy()
    df_calc["Saldo vs Demanda"] = df_calc[almox_dem] - df_calc["Demanda Pedido"]
    df_calc["Saldo Real"] = (df_calc[almox_dem]
        + (df_calc["Qtde Ordens Abertas"] if usar_ordens_flt and tem_ordens else 0)
        - df_calc["Demanda Pedido"] - df_calc["Qtde Pendente OP"])
    df_calc["Status"] = df_calc.apply(lambda r: calc_status(r, almox_dem, almox_seg, usar_ordens_flt), axis=1)

    if busca_flt:
        df_calc = df_calc[df_calc["Codigo"].astype(str).str.contains(busca_flt, case=False, na=False) |
                          df_calc["Descricao"].astype(str).str.contains(busca_flt, case=False, na=False)]
    if filtro_st_flt != "Todos":
        df_calc = df_calc[df_calc["Status"] == filtro_st_flt]

    colunas_exibir = ["Codigo","Descricao", almox_dem]
    if almox_seg != almox_dem: colunas_exibir.append(almox_seg)
    colunas_exibir += ["Demanda Pedido","Qtde Pendente OP"]
    if tem_ordens:    colunas_exibir.append("Qtde Ordens Abertas")
    colunas_exibir += ["Saldo vs Demanda","Saldo Real"]
    if tem_parametros: colunas_exibir.append("Estq Seg")
    colunas_exibir.append("Status")
    colunas_exibir = [c for c in colunas_exibir if c in df_calc.columns]

    st.markdown(f'<div style="font-size:10px;padding:4px 0"><b>{len(df_calc)}</b> item(ns) exibido(s)</div>', unsafe_allow_html=True)
    st.dataframe(df_calc[colunas_exibir].style.apply(estilo_linha, axis=1), use_container_width=True, height=400)

    col_d1, col_d2 = st.columns(2)
    with col_d1: st.download_button("📥 Baixar CSV",   df_calc[colunas_exibir].to_csv(index=False).encode("utf-8"), "pcp_filtrado.csv",  mime="text/csv",                                               key="dl_flt_csv")
    with col_d2: st.download_button("📥 Baixar Excel", exportar_excel_formatado(df_calc[colunas_exibir]),             "pcp_filtrado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_flt_xlsx")
    st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# ABA ESTOQUES
# ══════════════════════════════════════════════════════════════
with tab_estoques:
    st.markdown('<div class="panel"><div class="panel-header">Estoque de Segurança</div>', unsafe_allow_html=True)
    if not tem_parametros:
        st.markdown('<div class="msg-info">ℹ️ Importe o arquivo de Parâmetros para acessar análise de Estoque de Segurança.</div>', unsafe_allow_html=True)
    else:
        col_e1, col_e2, col_e3, col_e4, col_e5 = st.columns([1, 0.7, 0.7, 0.8, 1.2])
        with col_e1: almox_estq = st.selectbox("Almox", opcoes_almox, index=min(1,len(opcoes_almox)-1), key="estq_almox")
        with col_e2: inc_op_e  = st.checkbox("+ Pendente OP",     value=False, key="estq_inc_op")
        with col_e3: inc_ord_e = st.checkbox("+ Ordens Abertas",  value=False, key="estq_inc_ordens") if tem_ordens else False
        with col_e4: apenas_ab = st.checkbox("Só abaixo",         value=False, key="estq_apenas")
        with col_e5: busca_e   = st.text_input("Pesquisar...", key="busca_estq", placeholder="Código ou descrição...")

        df_estq = df_full[["Codigo","Descricao", almox_estq,"Demanda Pedido","Qtde Pendente OP","Estq Seg"]].copy()
        if tem_ordens: df_estq["Qtde Ordens Abertas"] = df_full["Qtde Ordens Abertas"]
        ef = df_estq[almox_estq].copy()
        if inc_op_e:  ef = ef + df_estq["Qtde Pendente OP"]
        if inc_ord_e and tem_ordens: ef = ef + df_estq["Qtde Ordens Abertas"]
        df_estq["Saldo Efetivo"] = ef
        df_estq["Diferença"] = df_estq["Saldo Efetivo"] - df_estq["Estq Seg"]
        df_estq["Status"]    = df_estq["Diferença"].apply(lambda x: "ABAIXO" if x < 0 else "OK")
        if apenas_ab: df_estq = df_estq[df_estq["Status"] == "ABAIXO"]
        if busca_e:   df_estq = df_estq[df_estq["Codigo"].astype(str).str.contains(busca_e, case=False, na=False) |
                                          df_estq["Descricao"].astype(str).str.contains(busca_e, case=False, na=False)]
        df_estq = df_estq.sort_values("Diferença")

        n_ab = (df_estq["Status"]=="ABAIXO").sum(); n_ok_e = (df_estq["Status"]=="OK").sum()
        st.markdown(f"""<div class="metric-row">
          {metric_box("⬇️ Abaixo", n_ab)}{metric_box("✅ OK", n_ok_e)}{metric_box("Almox", almox_estq)}
        </div>""", unsafe_allow_html=True)

        def estilo_estq(row):
            if row["Status"] == "ABAIXO": return ["background-color:#ffc7ce; color:#9c0006"] * len(row)
            return ["background-color:#f0fff4; color:#14532d"] * len(row)

        cols_estq = [c for c in ["Codigo","Descricao", almox_estq,"Saldo Efetivo","Estq Seg","Diferença","Status"] if c in df_estq.columns]
        st.markdown(f'<div style="font-size:10px;padding:4px 0"><b>{len(df_estq)}</b> item(ns)</div>', unsafe_allow_html=True)
        st.dataframe(df_estq[cols_estq].style.apply(estilo_estq, axis=1), use_container_width=True, height=400)

        col_d1, col_d2 = st.columns(2)
        with col_d1: st.download_button("📥 CSV",   df_estq[cols_estq].to_csv(index=False).encode("utf-8"), "estoques.csv",  mime="text/csv", key="dl_estq_csv")
        with col_d2: st.download_button("📥 Excel", exportar_excel_formatado(df_estq[cols_estq]),             "estoques.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_estq_xlsx")
    st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# ABA DEMANDAS
# ══════════════════════════════════════════════════════════════
with tab_demandas:
    st.markdown('<div class="panel"><div class="panel-header">Análise de Demandas</div>', unsafe_allow_html=True)

    col_d1, col_d2, col_d3, col_d4, col_d5, col_d6 = st.columns([1, 0.6, 0.6, 0.6, 0.7, 1.2])
    with col_d1: almox_dem2 = st.selectbox("Almox", opcoes_almox, index=0, key="dem_almox")
    with col_d2: inc_dc    = st.checkbox("DC",          value=True, key="dem_dc")
    with col_d3: inc_op_d  = st.checkbox("Pend. OP",    value=True, key="dem_op")
    with col_d4: inc_ord_d = st.checkbox("Ordens",      value=True, key="dem_ordens") if tem_ordens else False
    with col_d5: ap_dem    = st.checkbox("Só c/ demanda", value=False, key="dem_apenas")
    with col_d6: busca_d   = st.text_input("Pesquisar...", key="busca_dem", placeholder="Código ou descrição...")

    df_dem = df_full[["Codigo","Descricao", almox_dem2,"Demanda Pedido","Qtde Pendente OP"]].copy()
    if tem_ordens: df_dem["Qtde Ordens Abertas"] = df_full["Qtde Ordens Abertas"]
    dem_total = pd.Series([0.0]*len(df_dem), index=df_dem.index)
    if inc_dc:   dem_total = dem_total + df_dem["Demanda Pedido"]
    if inc_op_d: dem_total = dem_total + df_dem["Qtde Pendente OP"]
    if inc_ord_d and tem_ordens: dem_total = dem_total + df_dem["Qtde Ordens Abertas"]
    df_dem["Demanda Total"]    = dem_total
    df_dem["Saldo vs Demanda"] = df_dem[almox_dem2] - dem_total
    df_dem["Cobertura %"]      = (df_dem[almox_dem2] / dem_total.replace(0, float("nan")) * 100).fillna(0).round(1)
    df_dem["Status"]           = df_dem["Saldo vs Demanda"].apply(lambda x: "FALTA" if x < 0 else "OK")
    if ap_dem:   df_dem = df_dem[df_dem["Demanda Total"] > 0]
    if busca_d:  df_dem = df_dem[df_dem["Codigo"].astype(str).str.contains(busca_d, case=False, na=False) |
                                   df_dem["Descricao"].astype(str).str.contains(busca_d, case=False, na=False)]
    df_dem = df_dem.sort_values("Saldo vs Demanda")

    n_falta_d = (df_dem["Status"]=="FALTA").sum()
    st.markdown(f"""<div class="metric-row">
      {metric_box("📦 Total", len(df_dem))}{metric_box("🔴 Falta", n_falta_d)}
      {metric_box("Dem. Total", fmt_num(df_dem["Demanda Total"].sum()))}{metric_box("Saldo Total", fmt_num(df_dem[almox_dem2].sum()))}
    </div>""", unsafe_allow_html=True)

    def estilo_dem(row):
        if row["Status"] == "FALTA": return ["background-color:#ffc7ce; color:#9c0006"] * len(row)
        return ["background-color:#f0fff4; color:#14532d"] * len(row)

    cols_dem = ["Codigo","Descricao", almox_dem2,"Demanda Pedido"]
    if inc_op_d: cols_dem.append("Qtde Pendente OP")
    if inc_ord_d and tem_ordens: cols_dem.append("Qtde Ordens Abertas")
    cols_dem += ["Demanda Total","Saldo vs Demanda","Cobertura %","Status"]
    cols_dem = [c for c in cols_dem if c in df_dem.columns]

    st.markdown(f'<div style="font-size:10px;padding:4px 0"><b>{len(df_dem)}</b> item(ns)</div>', unsafe_allow_html=True)
    st.dataframe(df_dem[cols_dem].style.apply(estilo_dem, axis=1), use_container_width=True, height=400)

    col_dl1, col_dl2 = st.columns(2)
    with col_dl1: st.download_button("📥 CSV",   df_dem[cols_dem].to_csv(index=False).encode("utf-8"), "demandas.csv",  mime="text/csv", key="dl_dem_csv")
    with col_dl2: st.download_button("📥 Excel", exportar_excel_formatado(df_dem[cols_dem]),             "demandas.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_dem_xlsx")
    st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# ABA SUGESTÃO
# ══════════════════════════════════════════════════════════════
with tab_sugestao:
    st.markdown('<div class="panel"><div class="panel-header">Sugestão de Produção</div>', unsafe_allow_html=True)

    col_sg1, col_sg2, col_sg3 = st.columns([1.2, 0.8, 0.8])
    with col_sg1:
        idx_sug = opcoes_almox.index("Saldo Almox 3") if "Saldo Almox 3" in opcoes_almox else 0
        almox_sug = st.selectbox("Almox de referência", opcoes_almox, index=idx_sug, key="sug_almox")
    with col_sg2: inc_ord_sug = st.checkbox("+ Ordens no saldo",       value=True, key="sug_ordens") if tem_ordens else False
    with col_sg3: inc_estq_sug= st.checkbox("+ Est.Seg na sugestão",   value=True, key="sug_estq")   if tem_parametros else False

    almox_s_sug = opcoes_almox[-1] if not any("30" in x for x in opcoes_almox) else next(x for x in opcoes_almox if "30" in x)
    saldo_sug = df_full[almox_sug].copy()
    if inc_ord_sug and tem_ordens: saldo_sug = saldo_sug + df_full["Qtde Ordens Abertas"]
    df_sug = df_full.copy()
    df_sug["Saldo Real"]   = saldo_sug - df_sug["Demanda Pedido"] - df_sug["Qtde Pendente OP"]
    df_sug["Abaixo Estq Seg"] = df_sug[almox_sug] < df_sug["Estq Seg"]
    df_sug["Status"]       = df_sug.apply(lambda r: calc_status(r, almox_sug, almox_s_sug, inc_ord_sug), axis=1)
    df_sug_filt = df_sug[df_sug["Status"].isin(["FALTA","RISCO"])].copy()

    if not df_sug_filt.empty:
        df_sug_filt["Qtde Sugerida"] = df_sug_filt.apply(
            lambda r: max(r["Demanda Pedido"] + r["Qtde Pendente OP"] - r[almox_sug]
                          + (r["Estq Seg"] if inc_estq_sug and tem_parametros else 0), 0), axis=1)
        df_sug_filt = df_sug_filt[df_sug_filt["Qtde Sugerida"] > 0].sort_values(
            by=["Status","Qtde Sugerida"], ascending=[True, False])

        n_urgente = (df_sug_filt["Status"]=="FALTA").sum()
        n_atencao = (df_sug_filt["Status"]=="RISCO").sum()

        st.markdown(f"""<div class="metric-row">
          {metric_box("🔴 Urgente", n_urgente)}{metric_box("🟡 Atenção", n_atencao)}{metric_box("Total", len(df_sug_filt))}
        </div>""", unsafe_allow_html=True)

        # Lista de itens sugeridos
        sug_items_html = f'<div class="panel-header" style="text-align:left;padding:4px 8px">📋 {len(df_sug_filt)} ITEM(NS) PRECISAM SER PRODUZIDOS</div>'
        for _, row in df_sug_filt.head(15).iterrows():
            cls   = "sug-item-falta" if row["Status"]=="FALTA" else "sug-item-risco"
            emoji = "🔴" if row["Status"]=="FALTA" else "🟡"
            sug_items_html += f"""<div class="sug-item {cls}">
              <div>
                <div class="sug-cod">{emoji} {row['Codigo']}</div>
                <div style="font-size:8.5px;color:#555">{str(row['Descricao'])[:45]}</div>
              </div>
              <div class="sug-qtd">Prod: {fmt_num(row['Qtde Sugerida'])}</div>
            </div>"""
        if len(df_sug_filt) > 15:
            sug_items_html += f'<div style="font-size:9px;color:#777;padding:4px 8px">+ {len(df_sug_filt)-15} outros (veja o Excel)</div>'
        st.markdown(f'<div style="border:1px solid #aaa;margin-bottom:6px">{sug_items_html}</div>', unsafe_allow_html=True)

        with st.expander("📄 Ver tabela completa", expanded=False):
            cols_sug_t = [c for c in ["Codigo","Descricao", almox_sug,"Demanda Pedido","Qtde Pendente OP","Estq Seg","Qtde Sugerida","Status"] if c in df_sug_filt.columns]
            st.dataframe(df_sug_filt[cols_sug_t].style.apply(estilo_linha, axis=1), use_container_width=True)

        df_down_sug = df_sug_filt[[c for c in ["Codigo","Descricao", almox_sug,"Demanda Pedido","Qtde Pendente OP","Estq Seg","Qtde Sugerida","Status"] if c in df_sug_filt.columns]].copy()
        st.download_button("📥 Baixar Sugestão (Excel)", exportar_excel_formatado(df_down_sug), "sugestao_producao.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_sugestao")
    else:
        st.markdown('<div class="msg-ok">✅ Nenhum item necessita produção. Todos os saldos estão adequados.</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
