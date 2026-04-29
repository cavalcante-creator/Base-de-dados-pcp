<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Monitoramento PCP</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Segoe UI',Tahoma,sans-serif;font-size:11px;background:#d4d8dc;color:#1a1a2e;min-height:100vh}

  /* ---- HEADER ---- */
  .header{background:#c8ccce;border:1px solid #999;border-bottom:2px solid #888;padding:6px 14px;display:flex;align-items:center;justify-content:space-between;margin-bottom:6px}
  .logo{font-size:22px;font-weight:900;color:#c0392b;font-style:italic;letter-spacing:-1px}
  .header-title{font-size:15px;font-weight:700;color:#111;text-align:center;flex:1}
  .date-section{text-align:right;min-width:120px}
  .date-label{font-size:9px;color:#555;margin-bottom:2px;border-bottom:1px solid #aaa;padding-bottom:1px}
  .date-val{font-size:11px;font-weight:700;color:#1a1a2e;padding:2px 0}

  /* ---- LAYOUT ---- */
  .dash{padding:6px}
  .row{display:flex;gap:6px;margin-bottom:6px}
  .col{display:flex;flex-direction:column;gap:6px}

  /* ---- PANELS ---- */
  .panel{background:#e8eaec;border:1px solid #aaa;overflow:hidden}
  .panel-header{background:#c0c4c6;border-bottom:1px solid #999;padding:3px 8px;font-size:10px;font-weight:700;color:#111;text-align:center;letter-spacing:.3px}

  /* ---- TABELAS ---- */
  .tbl-wrap{overflow:auto;max-height:200px}
  table{width:100%;border-collapse:collapse;font-size:9.5px}
  th{background:#c0c4c6;border:1px solid #999;padding:3px 5px;font-weight:700;white-space:nowrap;position:sticky;top:0;z-index:1;text-align:left}
  td{border:1px solid #ccc;padding:2px 5px;vertical-align:middle;white-space:nowrap}
  tr:nth-child(even) td{background:#dfe2e5}
  tr:hover td{background:#cdd0d3}

  /* ---- STATUS BADGES ---- */
  .badge{display:inline-flex;align-items:center;gap:2px;padding:1px 5px;border-radius:2px;font-size:9px;font-weight:700}
  .badge-ok{background:#c6efce;color:#1e6b2e;border:1px solid #9dc09d}
  .badge-risco{background:#ffeb9c;color:#7a5000;border:1px solid #d4b800}
  .badge-falta{background:#ffc7ce;color:#9c0006;border:1px solid #e0808a}

  /* ---- LINHA COLORIDA ---- */
  .row-falta td{background:#ffc7ce !important}
  .row-risco td{background:#ffeb9c !important}
  .row-ok td{background:#f0fff4}

  /* ---- SALDO X PEDIDOS ---- */
  .saldo-panel{background:#e8eaec;border:1px solid #aaa}
  .legend-row{display:flex;align-items:center;gap:5px;padding:2px 6px;font-size:9.5px}
  .dot{width:10px;height:10px;border-radius:50%;flex-shrink:0;border:1px solid rgba(0,0,0,.2)}
  .dot-red{background:#e74c3c}
  .dot-yellow{background:#e6a817}
  .dot-green{background:#27ae60}
  .saldo-subtitle{font-size:8.5px;color:#555;padding:2px 6px 3px;border-top:1px solid #ccc;margin-top:2px}

  /* ---- MINI CHART ---- */
  .chart-area{padding:4px 6px 0;display:flex;align-items:flex-end;gap:2px;height:76px;border-bottom:1px solid #bbb}
  .bar-grp{display:flex;flex-direction:column;align-items:center;flex:1}
  .bars{display:flex;align-items:flex-end;gap:1px;height:68px}
  .bar{border-radius:1px 1px 0 0;min-width:5px}
  .bar-lbl{font-size:7px;color:#666;text-align:center;margin-top:2px;transform:rotate(-40deg);transform-origin:top center;width:18px;overflow:hidden}
  .chart-btns{display:flex;justify-content:center;gap:10px;padding:4px 0}
  .cbtn{width:13px;height:13px;border-radius:50%;border:2px solid rgba(0,0,0,.3);cursor:pointer}

  /* ---- UPLOAD ÁREA ---- */
  .upload-section{padding:8px}
  .upload-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px}
  .upload-box{background:#dde0e3;border:1.5px dashed #999;border-radius:4px;padding:8px;text-align:center;cursor:pointer;transition:.15s}
  .upload-box:hover{border-color:#555;background:#d0d4d8}
  .upload-box.loaded{border-color:#27ae60;background:#e8f5e9}
  .upload-box.loaded .up-icon{color:#27ae60}
  .up-icon{font-size:18px;color:#888;margin-bottom:3px}
  .up-label{font-size:10px;font-weight:700;color:#333;margin-bottom:2px}
  .up-sub{font-size:8.5px;color:#666}
  input[type=file]{display:none}

  /* ---- BOTÕES ---- */
  .btn{display:inline-flex;align-items:center;gap:4px;padding:4px 10px;border:1px solid #999;background:#d0d4d8;cursor:pointer;font-size:10px;font-weight:600;color:#222;border-radius:2px;transition:.1s}
  .btn:hover{background:#bbbfc3;border-color:#777}
  .btn-primary{background:#1a3a5c;color:white;border-color:#0d2740}
  .btn-primary:hover{background:#0d2740}
  .btn-danger{background:#c0392b;color:white;border-color:#922b21}
  .btn-sm{padding:2px 7px;font-size:9px}

  /* ---- ABAS NAV ---- */
  .nav-tabs{display:flex;gap:0;border-bottom:2px solid #999;margin-bottom:6px}
  .ntab{padding:4px 12px;background:#c8ccce;border:1px solid #999;border-bottom:none;cursor:pointer;font-size:10px;font-weight:600;color:#555;margin-right:2px;border-radius:2px 2px 0 0}
  .ntab.active{background:#e8eaec;color:#111;border-bottom:2px solid #e8eaec;margin-bottom:-2px}
  .ntab:hover:not(.active){background:#bbbfc3}

  /* ---- CARDS STATUS ---- */
  .status-cards{display:grid;grid-template-columns:repeat(4,1fr);gap:5px;padding:6px}
  .scard{padding:6px 8px;border-radius:3px;cursor:pointer;border:1px solid rgba(0,0,0,.1)}
  .scard-total{background:#1a3a5c;color:white}
  .scard-falta{background:#c0392b;color:white}
  .scard-risco{background:#d68910;color:white}
  .scard-ok{background:#1e8449;color:white}
  .scard-lbl{font-size:9px;font-weight:600;opacity:.85;text-transform:uppercase;letter-spacing:.5px}
  .scard-num{font-size:24px;font-weight:800;line-height:1}
  .scard-sub{font-size:8px;opacity:.75;margin-top:1px}

  /* ---- FILTROS ---- */
  .filter-bar{display:flex;gap:6px;padding:5px 6px;align-items:center;background:#dde0e3;border-bottom:1px solid #bbb;flex-wrap:wrap}
  .filter-bar input[type=text]{flex:1;min-width:120px;padding:3px 7px;border:1px solid #aaa;background:#f0f2f4;font-size:10px;border-radius:2px}
  .filter-bar select{padding:3px 6px;border:1px solid #aaa;background:#f0f2f4;font-size:10px;border-radius:2px}
  .filter-bar label{font-size:10px;font-weight:700;white-space:nowrap;display:flex;align-items:center;gap:4px}

  /* ---- MSG ---- */
  .msg-info{background:#d0e8f8;border:1px solid #7ab3d4;color:#0c4a6e;padding:6px 10px;font-size:10px;border-radius:2px;margin:6px}
  .msg-warn{background:#fff3cd;border:1px solid #e6a817;color:#7a5000;padding:6px 10px;font-size:10px;border-radius:2px;margin:6px}
  .msg-ok{background:#d4edda;border:1px solid #7dbb8e;color:#155724;padding:6px 10px;font-size:10px;border-radius:2px;margin:6px}

  /* ---- METRICS BAR ---- */
  .metrics-bar{display:flex;gap:8px;padding:5px 6px;background:#dde0e3;border-bottom:1px solid #ccc;flex-wrap:wrap}
  .metric-chip{background:#cdd0d3;border:1px solid #bbb;border-radius:2px;padding:4px 8px;min-width:80px}
  .metric-chip-lbl{font-size:9px;color:#555;font-weight:700}
  .metric-chip-val{font-size:14px;font-weight:800;color:#1a1a2e}

  /* ---- DOWNLOAD BAR ---- */
  .dl-bar{display:flex;gap:5px;padding:5px 6px;border-top:1px solid #ccc;background:#dde0e3}

  /* ---- SUGESTÃO BOX ---- */
  .sug-urgente{border-left:3px solid #c0392b}
  .sug-atencao{border-left:3px solid #d68910}

  /* ---- SCROLLBAR ---- */
  ::-webkit-scrollbar{width:6px;height:6px}
  ::-webkit-scrollbar-track{background:#d0d4d8}
  ::-webkit-scrollbar-thumb{background:#999;border-radius:3px}
  ::-webkit-scrollbar-thumb:hover{background:#666}

  .hidden{display:none}
  .flex-center{display:flex;align-items:center;justify-content:center;padding:20px;color:#777;font-size:11px}

  /* ---- SALDOS específico ---- */
  .saldo-cfg{display:grid;grid-template-columns:1fr 1fr;gap:6px;padding:6px}
  .saldo-metrics{display:flex;gap:8px;padding:5px 6px;background:#dde0e3;border-bottom:1px solid #ccc}
</style>
</head>
<body>
<div class="dash">

  <!-- HEADER -->
  <div class="header">
    <div class="logo">Colafix</div>
    <div class="header-title">Monitoramento PCP</div>
    <div class="date-section">
      <div class="date-label">Data Relatório ▼</div>
      <div class="date-val" id="dataAtual"></div>
    </div>
  </div>

  <!-- NAV PRINCIPAL -->
  <div class="nav-tabs">
    <div class="ntab active" onclick="showMain('upload')">⬆ Upload de Dados</div>
    <div class="ntab" onclick="showMain('resumo')">📊 Resumo</div>
    <div class="ntab" onclick="showMain('filtros')">🔍 Filtros</div>
    <div class="ntab" onclick="showMain('saldos')">💰 Saldos</div>
    <div class="ntab" onclick="showMain('estoques')">📦 Estoques</div>
    <div class="ntab" onclick="showMain('demandas')">📋 Demandas</div>
    <div class="ntab" onclick="showMain('sugestao')">🏭 Sugestão</div>
  </div>

  <!-- ===== ABA: UPLOAD ===== -->
  <div id="tab-upload">
    <div class="panel">
      <div class="panel-header">Upload e Processamento de Dados</div>
      <div class="upload-section">
        <div class="upload-grid">
          <div>
            <div class="upload-box" id="box-saldo-csv" onclick="document.getElementById('f-saldo-csv').click()">
              <div class="up-icon">📊</div>
              <div class="up-label">Saldo (CSV)</div>
              <div class="up-sub" id="sub-saldo-csv">Clique para selecionar</div>
            </div>
            <input type="file" id="f-saldo-csv" accept=".csv" onchange="carregarSaldoCSV(this)">
          </div>
          <div>
            <div class="upload-box" id="box-perfil" onclick="document.getElementById('f-perfil').click()">
              <div class="up-icon">📋</div>
              <div class="up-label">Perfil (CSV)</div>
              <div class="up-sub" id="sub-perfil">Clique para selecionar</div>
            </div>
            <input type="file" id="f-perfil" accept=".csv" onchange="carregarPerfil(this)">
          </div>
          <div>
            <div class="upload-box" id="box-ordens" onclick="document.getElementById('f-ordens').click()">
              <div class="up-icon">🏭</div>
              <div class="up-label">Ordens (CSV)</div>
              <div class="up-sub" id="sub-ordens">Clique para selecionar</div>
            </div>
            <input type="file" id="f-ordens" accept=".csv" onchange="carregarOrdens(this)">
          </div>
          <div>
            <div class="upload-box" id="box-previsao" onclick="document.getElementById('f-previsao').click()">
              <div class="up-icon">📈</div>
              <div class="up-label">Previsão (Excel)</div>
              <div class="up-sub" id="sub-previsao">Clique para selecionar</div>
            </div>
            <input type="file" id="f-previsao" accept=".xlsx,.xls" onchange="carregarPrevisao(this)">
          </div>
          <div>
            <div class="upload-box" id="box-params" onclick="document.getElementById('f-params').click()">
              <div class="up-icon">⚙️</div>
              <div class="up-label">Parâmetros (Excel)</div>
              <div class="up-sub" id="sub-params">Clique para selecionar</div>
            </div>
            <input type="file" id="f-params" accept=".xlsx,.xls" onchange="carregarParametros(this)">
          </div>
          <div>
            <div class="upload-box" id="box-saldo-pdf" onclick="document.getElementById('f-saldo-pdf').click()">
              <div class="up-icon">📄</div>
              <div class="up-label">Saldo (PDF — info)</div>
              <div class="up-sub" id="sub-saldo-pdf">Clique para selecionar</div>
            </div>
            <input type="file" id="f-saldo-pdf" accept=".pdf" onchange="infoSaldoPDF(this)">
          </div>
        </div>
        <div id="upload-msgs"></div>
        <div style="display:flex;gap:6px;margin-top:8px;padding:0 2px">
          <button class="btn btn-primary" onclick="gerarDashboard()">▶ Gerar Dashboard</button>
          <button class="btn btn-danger" onclick="limparBase()">🗑 Limpar Base</button>
        </div>
        <div id="base-preview" style="margin-top:8px"></div>
      </div>
    </div>
  </div>

  <!-- ===== ABA: RESUMO ===== -->
  <div id="tab-resumo" class="hidden">
    <div class="panel" style="margin-bottom:6px">
      <div class="panel-header">⚙️ Configurar Almoxarifados</div>
      <div class="filter-bar">
        <label>Almox — Demanda/Pedidos:</label>
        <select id="res-almox-dem" onchange="renderResumo()"></select>
        <label>Almox — Estoque de Segurança:</label>
        <select id="res-almox-seg" onchange="renderResumo()"></select>
        <label><input type="checkbox" id="res-ordens" onchange="renderResumo()" checked> Incluir Ordens Abertas</label>
      </div>
    </div>
    <div class="panel" style="margin-bottom:6px">
      <div class="panel-header">Resumo de Status</div>
      <div class="status-cards" id="statusCards"></div>
    </div>
    <div class="panel">
      <div class="panel-header">Visão Geral — clique nos cards para filtrar por status</div>
      <div class="tbl-wrap" style="max-height:380px">
        <table>
          <thead><tr><th>Codigo</th><th>Descrição</th><th id="res-col-almox">Saldo</th><th>Demanda Pedido</th><th>Saldo Real</th><th>Status</th></tr></thead>
          <tbody id="resumoBody"></tbody>
        </table>
      </div>
      <div class="dl-bar">
        <button class="btn btn-sm" onclick="downloadCSV('resumo')">📥 CSV</button>
        <button class="btn btn-sm" onclick="downloadXLSX('resumo')">📥 Excel</button>
      </div>
    </div>
  </div>

  <!-- ===== ABA: FILTROS ===== -->
  <div id="tab-filtros" class="hidden">
    <div class="panel">
      <div class="panel-header">Filtros e Análise Geral</div>
      <div class="filter-bar">
        <label>Almox Demanda:</label>
        <select id="flt-almox-dem" onchange="renderFiltros()"></select>
        <label>Almox Est.Seg:</label>
        <select id="flt-almox-seg" onchange="renderFiltros()"></select>
        <label><input type="checkbox" id="flt-ordens" onchange="renderFiltros()" checked> Incluir Ordens Abertas</label>
        <input type="text" id="busca-flt" placeholder="Pesquisar código/descrição..." oninput="renderFiltros()">
        <select id="flt-status" onchange="renderFiltros()">
          <option value="">Todos</option>
          <option value="FALTA">FALTA</option>
          <option value="RISCO">RISCO</option>
          <option value="OK">OK</option>
        </select>
        <select id="flt-ocultar" multiple style="max-height:40px;min-width:120px;font-size:9px" title="Ctrl+clique para ocultar itens">
        </select>
        <button class="btn btn-sm" onclick="aplicarOcultar()">Ocultar sel.</button>
        <button class="btn btn-sm" onclick="limparOcultar()">Limpar ocultos</button>
      </div>
      <div class="metrics-bar" id="flt-metrics"></div>
      <div class="tbl-wrap" style="max-height:420px">
        <table>
          <thead>
            <tr>
              <th>Codigo</th><th>Descrição</th><th id="flt-col-almox-d">Saldo Almox</th><th id="flt-col-almox-s">Est.Seg Almox</th>
              <th>Demanda Pedido</th><th>Qtde Pend. OP</th><th>Ordens Abertas</th>
              <th>Saldo vs Demanda</th><th>Saldo Real</th><th>Est.Seg</th><th>Status</th>
            </tr>
          </thead>
          <tbody id="filtrosBody"></tbody>
        </table>
      </div>
      <div class="dl-bar">
        <button class="btn btn-sm" onclick="downloadCSV('filtros')">📥 CSV</button>
        <button class="btn btn-sm" onclick="downloadXLSX('filtros')">📥 Excel</button>
      </div>
    </div>
  </div>

  <!-- ===== ABA: SALDOS ===== -->
  <div id="tab-saldos" class="hidden">
    <div class="panel">
      <div class="panel-header">💰 Saldos por Almoxarifado</div>
      <div class="filter-bar">
        <label>Almox de referência:</label>
        <select id="saldo-ref" onchange="renderSaldos()"></select>
        <label><input type="checkbox" id="saldo-mostrar-zero" onchange="renderSaldos()"> Mostrar saldo zero</label>
        <input type="text" id="busca-saldo" placeholder="Pesquisar item..." oninput="renderSaldos()">
      </div>
      <div class="metrics-bar" id="saldo-metrics"></div>
      <div class="tbl-wrap" style="max-height:420px">
        <table>
          <thead><tr id="saldo-thead"><th>Codigo</th><th>Descrição</th></tr></thead>
          <tbody id="saldoBody"></tbody>
        </table>
      </div>
      <div class="dl-bar">
        <button class="btn btn-sm" onclick="downloadCSV('saldos')">📥 CSV</button>
        <button class="btn btn-sm" onclick="downloadXLSX('saldos')">📥 Excel</button>
      </div>
    </div>
  </div>

  <!-- ===== ABA: ESTOQUES ===== -->
  <div id="tab-estoques" class="hidden">
    <div class="panel">
      <div class="panel-header">📦 Estoque de Segurança</div>
      <div class="filter-bar">
        <label>Almox:</label>
        <select id="estq-almox" onchange="renderEstoques()"></select>
        <label><input type="checkbox" id="estq-inc-op" onchange="renderEstoques()"> + Pendente OP</label>
        <label><input type="checkbox" id="estq-inc-ordens" onchange="renderEstoques()"> + Ordens Abertas</label>
        <label><input type="checkbox" id="estq-apenas-abaixo" onchange="renderEstoques()"> Só abaixo</label>
        <select id="estq-ordem" onchange="renderEstoques()">
          <option value="dif">Diferença (maior déficit)</option>
          <option value="cod">Código</option>
          <option value="status">Status</option>
        </select>
        <input type="text" id="busca-estq" placeholder="Pesquisar..." oninput="renderEstoques()">
      </div>
      <div class="metrics-bar" id="estq-metrics"></div>
      <div class="tbl-wrap" style="max-height:400px">
        <table>
          <thead><tr><th>Codigo</th><th>Descrição</th><th id="estq-col-almox">Saldo</th><th>Saldo Efetivo</th><th>Est.Seg</th><th>Diferença</th><th>Status</th></tr></thead>
          <tbody id="estqBody"></tbody>
        </table>
      </div>
      <div class="dl-bar">
        <button class="btn btn-sm" onclick="downloadCSV('estoques')">📥 CSV</button>
        <button class="btn btn-sm" onclick="downloadXLSX('estoques')">📥 Excel</button>
      </div>
    </div>
  </div>

  <!-- ===== ABA: DEMANDAS ===== -->
  <div id="tab-demandas" class="hidden">
    <div class="panel">
      <div class="panel-header">📋 Análise de Demandas</div>
      <div class="filter-bar">
        <label>Almox:</label>
        <select id="dem-almox" onchange="renderDemandas()"></select>
        <label><input type="checkbox" id="dem-dc" onchange="renderDemandas()" checked> DC</label>
        <label><input type="checkbox" id="dem-op" onchange="renderDemandas()" checked> Pendente OP</label>
        <label><input type="checkbox" id="dem-ordens" onchange="renderDemandas()" checked> Ordens</label>
        <label><input type="checkbox" id="dem-apenas" onchange="renderDemandas()"> Só com demanda</label>
        <input type="text" id="busca-dem" placeholder="Pesquisar..." oninput="renderDemandas()">
      </div>
      <div class="metrics-bar" id="dem-metrics"></div>
      <div class="tbl-wrap" style="max-height:400px">
        <table>
          <thead><tr><th>Codigo</th><th>Descrição</th><th id="dem-col-almox">Saldo</th><th>Demanda Pedido</th><th>Qtde Pend. OP</th><th>Ordens</th><th>Demanda Total</th><th>Saldo vs Dem.</th><th>Cobertura %</th><th>Status</th></tr></thead>
          <tbody id="demBody"></tbody>
        </table>
      </div>
      <div class="dl-bar">
        <button class="btn btn-sm" onclick="downloadCSV('demandas')">📥 CSV</button>
        <button class="btn btn-sm" onclick="downloadXLSX('demandas')">📥 Excel</button>
      </div>
    </div>
  </div>

  <!-- ===== ABA: SUGESTÃO ===== -->
  <div id="tab-sugestao" class="hidden">
    <div class="panel">
      <div class="panel-header">🏭 Sugestão de Produção</div>
      <div class="filter-bar">
        <label>Almox:</label>
        <select id="sug-almox" onchange="renderSugestao()"></select>
        <label><input type="checkbox" id="sug-ordens" onchange="renderSugestao()" checked> + Ordens no saldo</label>
        <label><input type="checkbox" id="sug-estq" onchange="renderSugestao()" checked> + Est.Seg na sugestão</label>
      </div>
      <div class="metrics-bar" id="sug-metrics"></div>

      <!-- Sugestão visual tipo card (replica bloco Python) -->
      <div id="sug-cards-box" style="padding:8px;background:#1e3a5c;margin:6px;border-radius:4px;color:white;display:none">
        <div style="font-size:12px;font-weight:700;margin-bottom:10px" id="sug-cards-titulo"></div>
        <div id="sug-cards-list"></div>
        <div id="sug-cards-mais" style="color:rgba(255,255,255,.6);font-size:11px;margin-top:6px"></div>
      </div>

      <div class="tbl-wrap" style="max-height:380px">
        <table>
          <thead><tr><th>Codigo</th><th>Descrição</th><th id="sug-col-almox">Saldo</th><th>Demanda Pedido</th><th>Qtde Pend. OP</th><th>Est.Seg</th><th>Qtde Sugerida</th><th>Status</th></tr></thead>
          <tbody id="sugBody"></tbody>
        </table>
      </div>
      <div class="dl-bar">
        <button class="btn btn-sm" onclick="downloadCSV('sugestao')">📥 CSV</button>
        <button class="btn btn-sm" onclick="downloadXLSX('sugestao')">📥 Excel</button>
      </div>
    </div>
  </div>

</div><!-- end dash -->

<script>
// ============================================================
// ESTADO GLOBAL
// ============================================================
const DB = {
  saldo: [],
  perfil: [],
  ordens: [],
  previsao: [],
  parametros: [],
};

let dfFull = [];
let opcaoAlmox = [];
let currentDownloadData = {};
let itensOcultados = new Set();
let filtroStatusGlobal = 'TODOS';

// ============================================================
// UTILIDADES
// ============================================================
function fmt(n, dec=0){
  if(n==null||isNaN(n)||n==='') return '';
  return Number(n).toLocaleString('pt-BR',{minimumFractionDigits:dec,maximumFractionDigits:dec});
}
function hoje(){
  return new Date().toLocaleDateString('pt-BR');
}
function tratarNumero(v){
  if(v==null||v===''||v===undefined) return 0;
  let s = String(v).trim().replace(/[^\d,.\-]/g,'');
  if(s.includes(',')) s = s.replace(/\./g,'').replace(',','.');
  return parseFloat(s)||0;
}
function msg(tipo, texto, container='upload-msgs'){
  const cls = tipo==='ok'?'msg-ok':tipo==='warn'?'msg-warn':'msg-info';
  document.getElementById(container).innerHTML += `<div class="${cls}">${texto}</div>`;
}
function clearMsgs(container='upload-msgs'){
  document.getElementById(container).innerHTML='';
}
function badgeHtml(st){
  if(st==='OK') return '<span class="badge badge-ok">✔ OK</span>';
  if(st==='RISCO') return '<span class="badge badge-risco">⚠ RISCO</span>';
  if(st==='FALTA') return '<span class="badge badge-falta">● FALTA</span>';
  if(st==='ABAIXO') return '<span class="badge badge-falta">⬇ ABAIXO</span>';
  return st;
}
function metricChip(lbl, val){
  return `<div class="metric-chip"><div class="metric-chip-lbl">${lbl}</div><div class="metric-chip-val">${val}</div></div>`;
}
function rowClass(st){
  if(st==='FALTA'||st==='ABAIXO') return 'row-falta';
  if(st==='RISCO') return 'row-risco';
  return '';
}

// ============================================================
// NAV
// ============================================================
const TABS = ['upload','resumo','filtros','saldos','estoques','demandas','sugestao'];
function showMain(tab){
  TABS.forEach(t=>{
    document.getElementById('tab-'+t).classList.toggle('hidden', t!==tab);
  });
  document.querySelectorAll('.ntab').forEach((el,i)=>{
    el.classList.toggle('active', TABS[i]===tab);
  });
  if(tab==='resumo') { populateAlmoxSelects(); renderResumo(); }
  if(tab==='filtros') { populateAlmoxSelects(); buildOcultarList(); renderFiltros(); }
  if(tab==='saldos') { populateAlmoxSelects(); renderSaldos(); }
  if(tab==='estoques') { populateAlmoxSelects(); renderEstoques(); }
  if(tab==='demandas') { populateAlmoxSelects(); renderDemandas(); }
  if(tab==='sugestao') { populateAlmoxSelects(); renderSugestao(); }
}

document.getElementById('dataAtual').textContent = hoje();

// ============================================================
// UPLOAD: SALDO CSV
// ============================================================
function carregarSaldoCSV(input){
  const file = input.files[0]; if(!file) return;
  Papa.parse(file,{header:true,skipEmptyLines:true,complete(res){
    const sample = res.data[0]||{};
    const keys = Object.keys(sample);
    DB.saldo = res.data.map(r=>{
      const cod = (r['Codigo']||r['CODIGO']||r['codigo']||'').trim();
      const obj = {Codigo: cod, 'Saldo Total': tratarNumero(r['Saldo Total']||r['SALDO TOTAL']||0)};
      // Capturar todas colunas de saldo almox dinamicamente
      keys.filter(k=>k.toUpperCase().includes('SALDO')||k.toUpperCase().includes('ALMOX')).forEach(k=>{ obj[k]=tratarNumero(r[k]); });
      return obj;
    }).filter(r=>r.Codigo);
    // Deduplica por codigo (mais recente)
    const seen={}; DB.saldo=DB.saldo.filter(r=>{ if(seen[r.Codigo]) return false; seen[r.Codigo]=true; return true; });
    setLoaded('box-saldo-csv','sub-saldo-csv',file.name,DB.saldo.length+' itens');
  }});
}

function infoSaldoPDF(input){
  const file = input.files[0]; if(!file) return;
  setLoaded('box-saldo-pdf','sub-saldo-pdf',file.name,'PDF recebido');
  msg('warn','⚠️ PDFs de saldo requerem processamento Python (pdfplumber). Use a aba Saldo CSV com arquivo exportado pelo sistema.');
}

// ============================================================
// UPLOAD: PERFIL CSV
// ============================================================
function carregarPerfil(input){
  const file = input.files[0]; if(!file) return;
  Papa.parse(file,{header:true,skipEmptyLines:true,complete(res){
    DB.perfil = res.data.map(r=>{
      const keys = Object.keys(r);
      const findKey = (...terms)=>keys.find(k=>terms.some(t=>k.toUpperCase().includes(t)))||'';
      const kItem=findKey('ITEM','COD','CODIGO');
      const kTipo=findKey('TIPO','TYPE');
      const kDt=findKey('DATA','DT');
      const kQtd=findKey('QUANT','QTD','QTDE');
      return {
        Item: (r[kItem]||'').trim(),
        Tipo: (r[kTipo]||'').trim().toUpperCase(),
        'Data Fim': r[kDt]||'',
        Quantidade: tratarNumero(r[kQtd]),
      };
    }).filter(r=>r.Item);
    setLoaded('box-perfil','sub-perfil',file.name,DB.perfil.length+' registros');
  }});
}

// ============================================================
// UPLOAD: ORDENS CSV
// ============================================================
function carregarOrdens(input){
  const file = input.files[0]; if(!file) return;
  Papa.parse(file,{header:true,skipEmptyLines:true,complete(res){
    DB.ordens = res.data;
    setLoaded('box-ordens','sub-ordens',file.name,res.data.length+' ordens');
  }});
}

// ============================================================
// UPLOAD: PREVISÃO EXCEL
// ============================================================
function carregarPrevisao(input){
  const file = input.files[0]; if(!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    const wb = XLSX.read(e.target.result,{type:'array'});
    const ws = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
    let hIdx = raw.findIndex(row=>row.some(c=>String(c).toUpperCase().includes('COD')));
    if(hIdx<0) hIdx=0;
    const headers = raw[hIdx].map(c=>String(c).toUpperCase().trim());
    const codIdx = headers.findIndex(h=>h.includes('COD'));
    const prodIdx = headers.findIndex(h=>h.includes('PROD')||h.includes('DESC'));
    DB.previsao = raw.slice(hIdx+1).filter(r=>r[codIdx]).map(r=>({
      COD: String(r[codIdx]).trim(),
      PRODUTO: String(prodIdx>=0?r[prodIdx]:'').trim()
    }));
    // Deduplica
    const seen={}; DB.previsao=DB.previsao.filter(r=>{ if(seen[r.COD]) return false; seen[r.COD]=true; return true; });
    setLoaded('box-previsao','sub-previsao',file.name,DB.previsao.length+' itens');
  };
  reader.readAsArrayBuffer(file);
}

// ============================================================
// UPLOAD: PARÂMETROS EXCEL
// ============================================================
function carregarParametros(input){
  const file = input.files[0]; if(!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    const wb = XLSX.read(e.target.result,{type:'array'});
    const ws = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws,{defval:0});
    DB.parametros = raw.map(r=>{
      const keys = Object.keys(r);
      const findKey = (...terms)=>keys.find(k=>terms.some(t=>k.toUpperCase().includes(t)))||'';
      const codKey=findKey('COD ITEM','CODIGO','CÓDIGO','COD');
      const estqKey=findKey('ESTQ SEG','ESTOQUE SEG','EST SEG','ESTSEG');
      const umKey=findKey(' UM','UNID','UM ');
      const consKey=findKey('CONS MEDIO','CONSUMO MED','CONS MED');
      const loteMinKey=findKey('LOTE MIN');
      const loteMultKey=findKey('LOTE MULT','MULT');
      return {
        Codigo: String(r[codKey]||'').trim(),
        'Estq Seg': tratarNumero(r[estqKey]),
        UM: String(r[umKey]||'').trim(),
        'CONS MEDIO': tratarNumero(r[consKey]),
        'LOTE MIN': tratarNumero(r[loteMinKey]),
        'LOTE MULT': tratarNumero(r[loteMultKey]),
      };
    }).filter(r=>r.Codigo);
    setLoaded('box-params','sub-params',file.name,DB.parametros.length+' itens');
  };
  reader.readAsArrayBuffer(file);
}

function setLoaded(boxId, subId, nome, info){
  document.getElementById(boxId).classList.add('loaded');
  document.getElementById(subId).textContent = nome.substring(0,22)+(nome.length>22?'...':'')+'  ('+info+')';
}

// ============================================================
// LIMPAR BASE
// ============================================================
function limparBase(){
  DB.saldo=[]; DB.perfil=[]; DB.ordens=[]; DB.previsao=[]; DB.parametros=[];
  dfFull=[]; itensOcultados.clear();
  ['box-saldo-csv','box-saldo-pdf','box-perfil','box-ordens','box-previsao','box-params'].forEach(id=>{
    const el=document.getElementById(id); if(el) el.classList.remove('loaded');
  });
  ['sub-saldo-csv','sub-saldo-pdf','sub-perfil','sub-ordens','sub-previsao','sub-params'].forEach(id=>{
    const el=document.getElementById(id); if(el) el.textContent='Clique para selecionar';
  });
  clearMsgs();
  document.getElementById('base-preview').innerHTML='';
  msg('ok','✅ Base de dados limpa com sucesso.');
}

// ============================================================
// CALCULAR dfFull
// ============================================================
function calcularBase(){
  clearMsgs();

  const base = DB.previsao.length>0
    ? DB.previsao.map(r=>({Codigo:r.COD, Descricao:r.PRODUTO}))
    : DB.saldo.map(r=>({Codigo:r.Codigo, Descricao:r.Codigo}));

  if(base.length===0){
    msg('warn','⚠️ Nenhum dado carregado. Faça upload dos arquivos primeiro.');
    return false;
  }

  const saldoMap = {};
  DB.saldo.forEach(r=>{ saldoMap[r.Codigo]=r; });

  // Detectar colunas de almox
  const primeiroSaldo = DB.saldo[0]||{};
  const allSaldoKeys = Object.keys(primeiroSaldo).filter(k=>k!=='Codigo');
  opcaoAlmox = allSaldoKeys.filter(k=>k.toUpperCase().includes('SALDO')||k.toUpperCase().includes('ALMOX'));
  if(opcaoAlmox.length===0) opcaoAlmox = allSaldoKeys;
  if(opcaoAlmox.length===0) opcaoAlmox = ['Saldo Total'];

  // DC: demanda pedido
  const dcMap = {};
  DB.perfil.filter(r=>r.Tipo==='DC').forEach(r=>{
    dcMap[r.Item]=(dcMap[r.Item]||0)+r.Quantidade;
  });

  // OP: pendente OP
  const opMap = {};
  DB.perfil.filter(r=>['OP','ORDEM','LIBERADA'].includes(r.Tipo)).forEach(r=>{
    opMap[r.Item]=(opMap[r.Item]||0)+r.Quantidade;
  });

  // Ordens abertas
  const ordensMap = {};
  if(DB.ordens.length>0){
    const firstRow = DB.ordens[0];
    const keys = Object.keys(firstRow);
    const kCod=keys.find(k=>['COD','ITEM','PRODUTO','CODIGO','CÓDIGO'].some(t=>k.toUpperCase().includes(t)));
    const kQtd=keys.find(k=>['QTDE','QTD','QUANTIDADE','SALDO ORDEM'].some(t=>k.toUpperCase().includes(t)));
    const kSt=keys.find(k=>['STATUS','SITUAÇÃO','SITUACAO','SIT'].some(t=>k.toUpperCase().includes(t)));
    const statusAberto=['ABERTO','ABERTA','EM ABERTO','LIBERADA','LIBERADO','PENDENTE','A PRODUZIR'];
    DB.ordens.forEach(r=>{
      if(kSt){
        const st=String(r[kSt]||'').toUpperCase().trim();
        if(!statusAberto.includes(st)) return;
      }
      const cod=String(r[kCod]||'').trim();
      ordensMap[cod]=(ordensMap[cod]||0)+tratarNumero(r[kQtd]);
    });
  }

  const paramMap = {};
  DB.parametros.forEach(r=>{ paramMap[r.Codigo]=r; });

  dfFull = base.map(row=>{
    const cod = row.Codigo;
    const sal = saldoMap[cod]||{};
    const pr = paramMap[cod]||{};
    const obj = {
      Codigo: cod,
      Descricao: row.Descricao,
      'Saldo Total': tratarNumero(sal['Saldo Total']),
      'Demanda Pedido': dcMap[cod]||0,
      'Qtde Pendente OP': opMap[cod]||0,
      'Qtde Ordens Abertas': ordensMap[cod]||0,
      'Estq Seg': tratarNumero(pr['Estq Seg']),
      UM: pr.UM||'',
      'CONS MEDIO': tratarNumero(pr['CONS MEDIO']),
      'LOTE MIN': tratarNumero(pr['LOTE MIN']),
      'LOTE MULT': tratarNumero(pr['LOTE MULT']),
    };
    opcaoAlmox.forEach(k=>{ obj[k]=tratarNumero(sal[k]); });
    return obj;
  });

  return true;
}

function calcStatus(row, almoxDem, almoxSeg, usarOrdens){
  const saldoReal = (row[almoxDem]||0) + (usarOrdens ? row['Qtde Ordens Abertas'] : 0)
    - row['Demanda Pedido'] - row['Qtde Pendente OP'];
  const abaixoSeg = (row[almoxSeg]||0) < row['Estq Seg'];
  if(saldoReal < 0) return 'FALTA';
  if(abaixoSeg) return 'RISCO';
  if(row['Demanda Pedido']+row['Qtde Pendente OP'] >= (row[almoxSeg]||0)*0.5 && (row['Demanda Pedido']+row['Qtde Pendente OP'])>0) return 'RISCO';
  return 'OK';
}

// ============================================================
// GERAR DASHBOARD
// ============================================================
function gerarDashboard(){
  if(!calcularBase()) return;
  populateAlmoxSelects();
  document.getElementById('base-preview').innerHTML =
    `<div class="msg-ok">✅ Base gerada: ${dfFull.length} itens | Almoxarifados: ${opcaoAlmox.join(', ')}</div>`;
  showMain('resumo');
}

function populateAlmoxSelects(){
  const selects=[
    {id:'res-almox-dem', pref:'3'},
    {id:'res-almox-seg', pref:'30'},
    {id:'flt-almox-dem', pref:'3'},
    {id:'flt-almox-seg', pref:'30'},
    {id:'saldo-ref', pref:''},
    {id:'estq-almox', pref:'30'},
    {id:'dem-almox', pref:'3'},
    {id:'sug-almox', pref:'3'},
  ];
  selects.forEach(({id, pref})=>{
    const el=document.getElementById(id); if(!el) return;
    const prev=el.value;
    el.innerHTML=opcaoAlmox.map(o=>`<option value="${o}">${o}</option>`).join('');
    const prefEl=opcaoAlmox.find(a=>pref&&a.includes(pref));
    if(prev && opcaoAlmox.includes(prev)) el.value=prev;
    else if(prefEl) el.value=prefEl;
  });
}

// ============================================================
// RESUMO
// ============================================================
function renderResumo(){
  if(!dfFull.length){ if(!calcularBase()) return; if(!dfFull.length) return; }
  const almD=(document.getElementById('res-almox-dem')||{}).value||opcaoAlmox[0];
  const almS=(document.getElementById('res-almox-seg')||{}).value||opcaoAlmox[opcaoAlmox.length-1]||almD;
  const usarOrdens=(document.getElementById('res-ordens')||{}).checked!==false;

  const rows = dfFull.map(r=>({...r,
    'Saldo Real': (r[almD]||0)+(usarOrdens?r['Qtde Ordens Abertas']:0)-r['Demanda Pedido']-r['Qtde Pendente OP'],
    Status: calcStatus(r, almD, almS, usarOrdens)
  }));

  const nTotal=rows.length, nFalta=rows.filter(r=>r.Status==='FALTA').length;
  const nRisco=rows.filter(r=>r.Status==='RISCO').length, nOk=rows.filter(r=>r.Status==='OK').length;

  document.getElementById('statusCards').innerHTML=`
    <div class="scard scard-total" onclick="filtrarEIr('TODOS')"><div class="scard-lbl">📦 Total</div><div class="scard-num">${nTotal}</div><div class="scard-sub">itens na base</div></div>
    <div class="scard scard-falta" onclick="filtrarEIr('FALTA')"><div class="scard-lbl">🔴 Falta</div><div class="scard-num">${nFalta}</div><div class="scard-sub">produção urgente</div></div>
    <div class="scard scard-risco" onclick="filtrarEIr('RISCO')"><div class="scard-lbl">🟡 Risco</div><div class="scard-num">${nRisco}</div><div class="scard-sub">abaixo est. seg.</div></div>
    <div class="scard scard-ok" onclick="filtrarEIr('OK')"><div class="scard-lbl">🟢 OK</div><div class="scard-num">${nOk}</div><div class="scard-sub">dentro do esperado</div></div>
  `;

  const colEl=document.getElementById('res-col-almox');
  if(colEl) colEl.textContent=almD;

  currentDownloadData.resumo = rows;
  document.getElementById('resumoBody').innerHTML = rows.map(r=>{
    const cls=rowClass(r.Status);
    return `<tr class="${cls}">
      <td><b>${r.Codigo}</b></td>
      <td style="max-width:200px;white-space:normal;font-size:9px">${r.Descricao}</td>
      <td style="text-align:right">${fmt(r[almD]||0)}</td>
      <td style="text-align:right">${fmt(r['Demanda Pedido'])}</td>
      <td style="text-align:right;font-weight:700;color:${r['Saldo Real']<0?'#9c0006':'#1e6b2e'}">${fmt(r['Saldo Real'])}</td>
      <td>${badgeHtml(r.Status)}</td>
    </tr>`;
  }).join('')||'<tr><td colspan="6" style="text-align:center;color:#888;padding:12px">Sem dados</td></tr>';
}

function filtrarEIr(status){
  filtroStatusGlobal=status;
  showMain('filtros');
  const el=document.getElementById('flt-status');
  if(el) el.value=(status==='TODOS'?'':status);
  renderFiltros();
}

// ============================================================
// FILTROS
// ============================================================
function buildOcultarList(){
  const sel=document.getElementById('flt-ocultar');
  if(!sel||!dfFull.length) return;
  const codigos=dfFull.map(r=>r.Codigo).sort();
  sel.innerHTML=codigos.map(c=>`<option value="${c}" ${itensOcultados.has(c)?'selected':''}>${c}</option>`).join('');
}
function aplicarOcultar(){
  const sel=document.getElementById('flt-ocultar');
  itensOcultados.clear();
  Array.from(sel.selectedOptions).forEach(o=>itensOcultados.add(o.value));
  renderFiltros();
}
function limparOcultar(){
  itensOcultados.clear();
  const sel=document.getElementById('flt-ocultar');
  if(sel) Array.from(sel.options).forEach(o=>o.selected=false);
  renderFiltros();
}

function renderFiltros(){
  if(!dfFull.length) return;
  const almD=(document.getElementById('flt-almox-dem')||{}).value||opcaoAlmox[0];
  const almS=(document.getElementById('flt-almox-seg')||{}).value||opcaoAlmox[opcaoAlmox.length-1]||almD;
  const usarOrdens=(document.getElementById('flt-ordens')||{}).checked!==false;
  const busca=(document.getElementById('busca-flt')||{}).value||'';
  const filtSt=(document.getElementById('flt-status')||{}).value||'';

  // Atualizar headers
  const hd=document.getElementById('flt-col-almox-d'); if(hd) hd.textContent='Saldo '+almD;
  const hs=document.getElementById('flt-col-almox-s'); if(hs) hs.textContent='Saldo '+almS;

  let rows=dfFull
    .filter(r=>!itensOcultados.has(r.Codigo))
    .map(r=>{
      const saldoReal=(r[almD]||0)+(usarOrdens?r['Qtde Ordens Abertas']:0)-r['Demanda Pedido']-r['Qtde Pendente OP'];
      const st=calcStatus(r, almD, almS, usarOrdens);
      return {...r, SaldoVsDem:(r[almD]||0)-r['Demanda Pedido'], SaldoReal:saldoReal, Status:st};
    });
  if(busca) rows=rows.filter(r=>r.Codigo.toLowerCase().includes(busca.toLowerCase())||String(r.Descricao).toLowerCase().includes(busca.toLowerCase()));
  if(filtSt) rows=rows.filter(r=>r.Status===filtSt);

  currentDownloadData.filtros=rows;
  const nFalta=rows.filter(r=>r.Status==='FALTA').length;
  const nRisco=rows.filter(r=>r.Status==='RISCO').length;
  const nOk=rows.filter(r=>r.Status==='OK').length;
  document.getElementById('flt-metrics').innerHTML=
    metricChip('Total',rows.length)+metricChip('🔴 Falta',nFalta)+metricChip('🟡 Risco',nRisco)+metricChip('🟢 OK',nOk)+
    metricChip('Almox Dem.',almD)+metricChip('Almox Seg.',almS)+
    (itensOcultados.size?metricChip('🚫 Ocultos',itensOcultados.size):'');

  document.getElementById('filtrosBody').innerHTML=rows.map(r=>{
    const cls=rowClass(r.Status);
    return `<tr class="${cls}">
      <td><b>${r.Codigo}</b></td>
      <td style="font-size:9px;max-width:160px;white-space:normal">${r.Descricao}</td>
      <td style="text-align:right">${fmt(r[almD]||0)}</td>
      <td style="text-align:right">${fmt(r[almS]||0)}</td>
      <td style="text-align:right">${fmt(r['Demanda Pedido'])}</td>
      <td style="text-align:right">${fmt(r['Qtde Pendente OP'])}</td>
      <td style="text-align:right">${fmt(r['Qtde Ordens Abertas'])}</td>
      <td style="text-align:right;font-weight:700;color:${r.SaldoVsDem<0?'#9c0006':'#1e6b2e'}">${fmt(r.SaldoVsDem)}</td>
      <td style="text-align:right;font-weight:700;color:${r.SaldoReal<0?'#9c0006':'#1e6b2e'}">${fmt(r.SaldoReal)}</td>
      <td style="text-align:right">${fmt(r['Estq Seg'])}</td>
      <td>${badgeHtml(r.Status)}</td>
    </tr>`;
  }).join('')||'<tr><td colspan="11" style="text-align:center;color:#888;padding:12px">Sem dados</td></tr>';
}

// ============================================================
// SALDOS
// ============================================================
function renderSaldos(){
  if(!dfFull.length) return;
  const almoxRef=(document.getElementById('saldo-ref')||{}).value||opcaoAlmox[0];
  const mostrarZero=(document.getElementById('saldo-mostrar-zero')||{}).checked;
  const busca=(document.getElementById('busca-saldo')||{}).value||'';

  // Monta join base + saldo
  const saldoMap={};
  DB.saldo.forEach(r=>{ saldoMap[r.Codigo]=r; });
  const base=DB.previsao.length>0
    ? DB.previsao.map(r=>({Codigo:r.COD,Descricao:r.PRODUTO}))
    : DB.saldo.map(r=>({Codigo:r.Codigo,Descricao:r.Codigo}));

  let rows=base.map(row=>{
    const sal=saldoMap[row.Codigo]||{};
    const obj={Codigo:row.Codigo,Descricao:row.Descricao};
    opcaoAlmox.forEach(k=>{ obj[k]=tratarNumero(sal[k]); });
    return obj;
  });

  if(!mostrarZero) rows=rows.filter(r=>(r[almoxRef]||0)>0);
  if(busca) rows=rows.filter(r=>r.Codigo.toLowerCase().includes(busca.toLowerCase())||String(r.Descricao).toLowerCase().includes(busca.toLowerCase()));

  currentDownloadData.saldos=rows;

  const total=rows.filter(r=>(r[almoxRef]||0)>0).length;
  const semSaldo=rows.filter(r=>(r[almoxRef]||0)===0).length;
  const somaRef=rows.reduce((s,r)=>s+(r[almoxRef]||0),0);
  document.getElementById('saldo-metrics').innerHTML=
    metricChip('Com saldo',total)+metricChip(`Soma ${almoxRef}`,fmt(somaRef))+metricChip('Sem saldo',semSaldo);

  // Build thead dinamico
  const thead=document.getElementById('saldo-thead');
  if(thead) thead.innerHTML='<th>Codigo</th><th>Descrição</th>'+opcaoAlmox.map(k=>`<th>${k}</th>`).join('');

  document.getElementById('saldoBody').innerHTML=rows.map(r=>`
    <tr>
      <td><b>${r.Codigo}</b></td>
      <td style="font-size:9px;max-width:180px;white-space:normal">${r.Descricao}</td>
      ${opcaoAlmox.map(k=>`<td style="text-align:right">${fmt(r[k]||0)}</td>`).join('')}
    </tr>`).join('')||'<tr><td colspan="10" style="text-align:center;color:#888;padding:12px">Sem dados</td></tr>';
}

// ============================================================
// ESTOQUES
// ============================================================
function renderEstoques(){
  if(!dfFull.length) return;
  const almox=(document.getElementById('estq-almox')||{}).value||opcaoAlmox[0];
  const incOP=(document.getElementById('estq-inc-op')||{}).checked;
  const incOrdens=(document.getElementById('estq-inc-ordens')||{}).checked;
  const apenasAbaixo=(document.getElementById('estq-apenas-abaixo')||{}).checked;
  const ordem=(document.getElementById('estq-ordem')||{}).value||'dif';
  const busca=(document.getElementById('busca-estq')||{}).value||'';

  const hcol=document.getElementById('estq-col-almox'); if(hcol) hcol.textContent=almox;

  let rows=dfFull.map(r=>{
    let ef=(r[almox]||0);
    if(incOP) ef+=r['Qtde Pendente OP'];
    if(incOrdens) ef+=r['Qtde Ordens Abertas'];
    const diff=ef-r['Estq Seg'];
    return {...r, SaldoEfetivo:ef, Diferenca:diff, StatusEstq:diff<0?'ABAIXO':'OK'};
  });
  if(apenasAbaixo) rows=rows.filter(r=>r.StatusEstq==='ABAIXO');
  if(busca) rows=rows.filter(r=>r.Codigo.toLowerCase().includes(busca.toLowerCase())||String(r.Descricao).toLowerCase().includes(busca.toLowerCase()));
  if(ordem==='dif') rows.sort((a,b)=>a.Diferenca-b.Diferenca);
  else if(ordem==='cod') rows.sort((a,b)=>a.Codigo.localeCompare(b.Codigo));
  else rows.sort((a,b)=>a.StatusEstq.localeCompare(b.StatusEstq));

  currentDownloadData.estoques=rows;
  const nAbaixo=rows.filter(r=>r.StatusEstq==='ABAIXO').length;
  const nOk=rows.filter(r=>r.StatusEstq==='OK').length;
  document.getElementById('estq-metrics').innerHTML=
    metricChip('⬇️ Abaixo',nAbaixo)+metricChip('✅ OK',nOk)+metricChip('Almox',almox)+metricChip('Total',rows.length);

  document.getElementById('estqBody').innerHTML=rows.map(r=>{
    const cls=r.StatusEstq==='ABAIXO'?'row-falta':'row-ok';
    return `<tr class="${cls}">
      <td><b>${r.Codigo}</b></td>
      <td style="font-size:9px;max-width:160px;white-space:normal">${r.Descricao}</td>
      <td style="text-align:right">${fmt(r[almox]||0)}</td>
      <td style="text-align:right">${fmt(r.SaldoEfetivo)}</td>
      <td style="text-align:right">${fmt(r['Estq Seg'])}</td>
      <td style="text-align:right;color:${r.Diferenca<0?'#9c0006':'#1e6b2e'};font-weight:700">${fmt(r.Diferenca)}</td>
      <td>${badgeHtml(r.StatusEstq)}</td>
    </tr>`;
  }).join('')||'<tr><td colspan="7" style="text-align:center;color:#888;padding:12px">Sem dados</td></tr>';
}

// ============================================================
// DEMANDAS
// ============================================================
function renderDemandas(){
  if(!dfFull.length) return;
  const almox=(document.getElementById('dem-almox')||{}).value||opcaoAlmox[0];
  const incDC=(document.getElementById('dem-dc')||{}).checked!==false;
  const incOP=(document.getElementById('dem-op')||{}).checked!==false;
  const incOrdens=(document.getElementById('dem-ordens')||{}).checked!==false;
  const apenasComDem=(document.getElementById('dem-apenas')||{}).checked;
  const busca=(document.getElementById('busca-dem')||{}).value||'';

  const hcol=document.getElementById('dem-col-almox'); if(hcol) hcol.textContent=almox;

  let rows=dfFull.map(r=>{
    let dem=0;
    if(incDC) dem+=r['Demanda Pedido'];
    if(incOP) dem+=r['Qtde Pendente OP'];
    if(incOrdens) dem+=r['Qtde Ordens Abertas'];
    const svd=(r[almox]||0)-dem;
    const cob=dem>0?Math.round((r[almox]||0)/dem*100):0;
    return {...r, DemandaTotal:dem, SaldoVsDem:svd, Cobertura:cob, Status:svd<0?'FALTA':'OK'};
  });
  if(apenasComDem) rows=rows.filter(r=>r.DemandaTotal>0);
  if(busca) rows=rows.filter(r=>r.Codigo.toLowerCase().includes(busca.toLowerCase())||String(r.Descricao).toLowerCase().includes(busca.toLowerCase()));
  rows.sort((a,b)=>a.SaldoVsDem-b.SaldoVsDem);
  currentDownloadData.demandas=rows;

  const nFalta=rows.filter(r=>r.Status==='FALTA').length;
  document.getElementById('dem-metrics').innerHTML=
    metricChip('📦 Total',rows.length)+metricChip('🔴 Falta',nFalta)+
    metricChip('Dem. Total',fmt(rows.reduce((s,r)=>s+r.DemandaTotal,0)))+
    metricChip('Saldo Total',fmt(rows.reduce((s,r)=>s+(r[almox]||0),0)));

  document.getElementById('demBody').innerHTML=rows.map(r=>{
    const cls=r.Status==='FALTA'?'row-falta':'';
    return `<tr class="${cls}">
      <td><b>${r.Codigo}</b></td>
      <td style="font-size:9px;max-width:140px;white-space:normal">${r.Descricao}</td>
      <td style="text-align:right">${fmt(r[almox]||0)}</td>
      <td style="text-align:right">${fmt(r['Demanda Pedido'])}</td>
      <td style="text-align:right">${fmt(r['Qtde Pendente OP'])}</td>
      <td style="text-align:right">${fmt(r['Qtde Ordens Abertas'])}</td>
      <td style="text-align:right">${fmt(r.DemandaTotal)}</td>
      <td style="text-align:right;color:${r.SaldoVsDem<0?'#9c0006':'#1e6b2e'};font-weight:700">${fmt(r.SaldoVsDem)}</td>
      <td style="text-align:right">${r.Cobertura}%</td>
      <td>${badgeHtml(r.Status)}</td>
    </tr>`;
  }).join('')||'<tr><td colspan="10" style="text-align:center;color:#888;padding:12px">Sem dados</td></tr>';
}

// ============================================================
// SUGESTÃO
// ============================================================
function renderSugestao(){
  if(!dfFull.length) return;
  const almox=(document.getElementById('sug-almox')||{}).value||opcaoAlmox[0];
  const incOrdens=(document.getElementById('sug-ordens')||{}).checked!==false;
  const incEstqSeg=(document.getElementById('sug-estq')||{}).checked!==false;
  const almS=opcaoAlmox.find(a=>a.includes('30'))||almox;

  const hcol=document.getElementById('sug-col-almox'); if(hcol) hcol.textContent=almox;

  let rows=dfFull.map(r=>{
    const saldo=(r[almox]||0)+(incOrdens?r['Qtde Ordens Abertas']:0);
    const saldoReal=saldo-r['Demanda Pedido']-r['Qtde Pendente OP'];
    const abaixoSeg=(r[almox]||0)<r['Estq Seg'];
    let st='OK';
    if(saldoReal<0) st='FALTA';
    else if(abaixoSeg) st='RISCO';
    else if(r['Demanda Pedido']+r['Qtde Pendente OP']>=(r[almS]||0)*0.5 && (r['Demanda Pedido']+r['Qtde Pendente OP'])>0) st='RISCO';
    const qtdSug=Math.max(r['Demanda Pedido']+r['Qtde Pendente OP']-(r[almox]||0)+(incEstqSeg?r['Estq Seg']:0),0);
    return {...r, SaldoReal:saldoReal, Status:st, QtdSug:qtdSug};
  }).filter(r=>r.Status!=='OK'&&r.QtdSug>0)
    .sort((a,b)=>(a.Status==='FALTA'?0:1)-(b.Status==='FALTA'?0:1)||b.QtdSug-a.QtdSug);

  currentDownloadData.sugestao=rows;
  const nUrgente=rows.filter(r=>r.Status==='FALTA').length;
  const nAtencao=rows.filter(r=>r.Status==='RISCO').length;
  document.getElementById('sug-metrics').innerHTML=
    metricChip('🔴 Urgente',nUrgente)+metricChip('🟡 Atenção',nAtencao)+metricChip('Total',rows.length);

  // Cards visuais tipo Python
  const cardsBox=document.getElementById('sug-cards-box');
  const cardsList=document.getElementById('sug-cards-list');
  const cardsMais=document.getElementById('sug-cards-mais');
  if(rows.length>0){
    cardsBox.style.display='block';
    document.getElementById('sug-cards-titulo').textContent=`📋 ${rows.length} ITEM(NS) PRECISAM SER PRODUZIDOS`;
    const top10=rows.slice(0,10);
    cardsList.innerHTML=top10.map(r=>{
      const cls=r.Status==='FALTA'?'sug-urgente':'sug-atencao';
      const emoji=r.Status==='FALTA'?'🔴':'🟡';
      return `<div class="${cls}" style="background:rgba(255,255,255,.12);border-radius:4px;padding:6px 10px;margin-bottom:5px;display:flex;justify-content:space-between;align-items:center;border:1px solid rgba(255,255,255,.1)">
        <div><div style="font-weight:600;font-size:11px">${emoji} ${r.Codigo}</div><div style="font-size:9px;opacity:.75">${String(r.Descricao||'').substring(0,50)}</div></div>
        <div style="background:rgba(255,255,255,.2);border-radius:3px;padding:3px 8px;font-weight:700;font-size:12px">Prod: ${fmt(r.QtdSug)}</div>
      </div>`;
    }).join('');
    cardsMais.textContent=rows.length>10?`+ ${rows.length-10} outros (baixe o Excel para ver todos)`:'';
  } else {
    cardsBox.style.display='none';
  }

  document.getElementById('sugBody').innerHTML=rows.length===0
    ?'<tr><td colspan="8" style="text-align:center;color:#27ae60;padding:12px">✅ Nenhum item necessita produção</td></tr>'
    :rows.map(r=>{
      const cls=r.Status==='FALTA'?'row-falta':'row-risco';
      return `<tr class="${cls}">
        <td><b>${r.Codigo}</b></td>
        <td style="font-size:9px;max-width:140px;white-space:normal">${r.Descricao}</td>
        <td style="text-align:right">${fmt(r[almox]||0)}</td>
        <td style="text-align:right">${fmt(r['Demanda Pedido'])}</td>
        <td style="text-align:right">${fmt(r['Qtde Pendente OP'])}</td>
        <td style="text-align:right">${fmt(r['Estq Seg'])}</td>
        <td style="text-align:right;font-weight:700;color:#1a3a5c">${fmt(r.QtdSug)}</td>
        <td>${badgeHtml(r.Status)}</td>
      </tr>`;
    }).join('');
}

// ============================================================
// DOWNLOADS
// ============================================================
function downloadCSV(key){
  const rows=currentDownloadData[key];
  if(!rows||!rows.length) return alert('Sem dados para exportar.');
  const cols=Object.keys(rows[0]);
  const csv=cols.join(',')+'\\n'+rows.map(r=>cols.map(c=>'"'+(r[c]??'')+'"').join(',')).join('\\n');
  const a=document.createElement('a');
  a.href='data:text/csv;charset=utf-8,\uFEFF'+encodeURIComponent(csv);
  a.download=key+'_pcp.csv'; a.click();
}
function downloadXLSX(key){
  const rows=currentDownloadData[key];
  if(!rows||!rows.length) return alert('Sem dados para exportar.');
  const ws=XLSX.utils.json_to_sheet(rows);
  // Formatar cabeçalho
  const range=XLSX.utils.decode_range(ws['!ref']||'A1');
  for(let c=range.s.c;c<=range.e.c;c++){
    const addr=XLSX.utils.encode_cell({r:0,c});
    if(!ws[addr]) continue;
    ws[addr].s={fill:{fgColor:{rgb:'1F4E78'}},font:{color:{rgb:'FFFFFF'},bold:true}};
  }
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Dados');
  XLSX.writeFile(wb,key+'_pcp.xlsx');
}
</script>
</body>
</html>
