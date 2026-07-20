<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard · Generador de Informes</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500&display=swap" rel="stylesheet">
<style>
  :root {
    --bg: #0a0a0f;
    --surface: #13131a;
    --surface2: #1c1c28;
    --border: #2a2a3d;
    --accent: #6c63ff;
    --accent2: #ff6584;
    --accent3: #43e97b;
    --accent4: #f7971e;
    --text: #e8e8f0;
    --text2: #8888aa;
    --text3: #5555770;
    --radius: 12px;
  }

  * { margin: 0; padding: 0; box-sizing: border-box; }

  body {
    background: var(--bg);
    color: var(--text);
    font-family: 'DM Sans', sans-serif;
    min-height: 100vh;
    overflow-x: hidden;
  }

  /* Animated background */
  body::before {
    content: '';
    position: fixed;
    top: -50%;
    left: -50%;
    width: 200%;
    height: 200%;
    background: radial-gradient(ellipse at 20% 20%, rgba(108,99,255,0.08) 0%, transparent 50%),
                radial-gradient(ellipse at 80% 80%, rgba(255,101,132,0.06) 0%, transparent 50%),
                radial-gradient(ellipse at 50% 50%, rgba(67,233,123,0.04) 0%, transparent 60%);
    animation: bgPulse 12s ease-in-out infinite alternate;
    pointer-events: none;
    z-index: 0;
  }
  @keyframes bgPulse {
    0% { transform: scale(1) rotate(0deg); }
    100% { transform: scale(1.1) rotate(5deg); }
  }

  .app {
    position: relative;
    z-index: 1;
    max-width: 1400px;
    margin: 0 auto;
    padding: 0 24px 60px;
  }

  /* HEADER */
  header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 28px 0 20px;
    border-bottom: 1px solid var(--border);
    margin-bottom: 36px;
  }

  .logo {
    display: flex;
    align-items: center;
    gap: 12px;
  }

  .logo-icon {
    width: 38px; height: 38px;
    background: linear-gradient(135deg, var(--accent), var(--accent2));
    border-radius: 10px;
    display: flex; align-items: center; justify-content: center;
    font-size: 18px;
  }

  .logo-text {
    font-family: 'Syne', sans-serif;
    font-size: 20px;
    font-weight: 800;
    letter-spacing: -0.5px;
  }
  .logo-text span { color: var(--accent); }

  .header-badge {
    background: var(--surface2);
    border: 1px solid var(--border);
    padding: 6px 14px;
    border-radius: 20px;
    font-size: 12px;
    color: var(--text2);
    letter-spacing: 0.5px;
  }

  /* UPLOAD ZONE */
  .upload-section {
    margin-bottom: 40px;
  }

  .section-label {
    font-family: 'Syne', sans-serif;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: var(--accent);
    margin-bottom: 14px;
  }

  .upload-zone {
    border: 2px dashed var(--border);
    border-radius: var(--radius);
    padding: 48px 32px;
    text-align: center;
    cursor: pointer;
    transition: all 0.3s ease;
    background: var(--surface);
    position: relative;
    overflow: hidden;
  }

  .upload-zone::before {
    content: '';
    position: absolute;
    inset: 0;
    background: linear-gradient(135deg, rgba(108,99,255,0.05), rgba(255,101,132,0.03));
    opacity: 0;
    transition: opacity 0.3s;
  }

  .upload-zone:hover, .upload-zone.drag-over {
    border-color: var(--accent);
    transform: translateY(-2px);
    box-shadow: 0 8px 32px rgba(108,99,255,0.2);
  }
  .upload-zone:hover::before, .upload-zone.drag-over::before { opacity: 1; }

  .upload-icon {
    font-size: 48px;
    margin-bottom: 16px;
    display: block;
    animation: float 3s ease-in-out infinite;
  }
  @keyframes float {
    0%, 100% { transform: translateY(0); }
    50% { transform: translateY(-8px); }
  }

  .upload-title {
    font-family: 'Syne', sans-serif;
    font-size: 18px;
    font-weight: 700;
    margin-bottom: 8px;
  }

  .upload-sub {
    color: var(--text2);
    font-size: 14px;
    margin-bottom: 20px;
  }

  .upload-btn {
    display: inline-block;
    background: linear-gradient(135deg, var(--accent), #8b7ff5);
    color: white;
    padding: 10px 24px;
    border-radius: 8px;
    font-size: 14px;
    font-weight: 500;
    cursor: pointer;
    transition: all 0.2s;
    border: none;
  }
  .upload-btn:hover { opacity: 0.9; transform: scale(1.02); }

  #fileInput { display: none; }

  /* FILE INFO */
  .file-info {
    display: none;
    align-items: center;
    gap: 12px;
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 14px 18px;
    margin-top: 14px;
  }
  .file-info.visible { display: flex; }
  .file-dot {
    width: 10px; height: 10px;
    background: var(--accent3);
    border-radius: 50%;
    flex-shrink: 0;
    box-shadow: 0 0 8px var(--accent3);
    animation: pulse 2s infinite;
  }
  @keyframes pulse { 0%,100%{opacity:1}50%{opacity:0.5} }
  .file-name { font-size: 14px; font-weight: 500; }
  .file-meta { font-size: 12px; color: var(--text2); margin-left: auto; }

  /* SHEET SELECTOR */
  .sheet-selector {
    display: none;
    gap: 8px;
    flex-wrap: wrap;
    margin-top: 16px;
  }
  .sheet-selector.visible { display: flex; }

  .sheet-tab {
    padding: 6px 16px;
    border-radius: 20px;
    font-size: 13px;
    cursor: pointer;
    border: 1px solid var(--border);
    background: var(--surface);
    color: var(--text2);
    transition: all 0.2s;
  }
  .sheet-tab.active {
    background: var(--accent);
    border-color: var(--accent);
    color: white;
  }
  .sheet-tab:hover:not(.active) { border-color: var(--accent); color: var(--text); }

  /* KPI CARDS */
  .kpi-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 16px;
    margin-bottom: 36px;
  }

  .kpi-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 20px;
    position: relative;
    overflow: hidden;
    transition: all 0.3s;
    animation: slideUp 0.4s ease forwards;
    opacity: 0;
    transform: translateY(20px);
  }
  .kpi-card:nth-child(1) { animation-delay: 0.05s; }
  .kpi-card:nth-child(2) { animation-delay: 0.1s; }
  .kpi-card:nth-child(3) { animation-delay: 0.15s; }
  .kpi-card:nth-child(4) { animation-delay: 0.2s; }
  @keyframes slideUp { to { opacity:1; transform:translateY(0); } }

  .kpi-card::after {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
  }
  .kpi-card:nth-child(1)::after { background: var(--accent); }
  .kpi-card:nth-child(2)::after { background: var(--accent2); }
  .kpi-card:nth-child(3)::after { background: var(--accent3); }
  .kpi-card:nth-child(4)::after { background: var(--accent4); }

  .kpi-card:hover { transform: translateY(-4px); box-shadow: 0 12px 32px rgba(0,0,0,0.4); }

  .kpi-label { font-size: 11px; color: var(--text2); text-transform: uppercase; letter-spacing: 1px; margin-bottom: 10px; }
  .kpi-value { font-family: 'Syne', sans-serif; font-size: 28px; font-weight: 800; }
  .kpi-sub { font-size: 12px; color: var(--text2); margin-top: 4px; }

  /* CHARTS + TABLE LAYOUT */
  .dashboard-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 20px;
    margin-bottom: 28px;
  }

  .panel {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 22px;
    animation: slideUp 0.5s ease forwards;
    opacity: 0;
    transform: translateY(20px);
  }
  .panel:nth-child(1) { animation-delay: 0.2s; }
  .panel:nth-child(2) { animation-delay: 0.3s; }
  .panel.full-width { grid-column: 1 / -1; }

  .panel-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 18px;
  }

  .panel-title {
    font-family: 'Syne', sans-serif;
    font-size: 15px;
    font-weight: 700;
  }

  .panel-badge {
    font-size: 11px;
    background: var(--surface2);
    color: var(--text2);
    padding: 4px 10px;
    border-radius: 10px;
    border: 1px solid var(--border);
  }

  .chart-wrap {
    position: relative;
    height: 280px;
  }

  /* TABLE */
  .table-scroll {
    overflow-x: auto;
    border-radius: 8px;
  }

  table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
  }

  thead th {
    background: var(--surface2);
    color: var(--text2);
    font-family: 'Syne', sans-serif;
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    padding: 12px 14px;
    text-align: left;
    white-space: nowrap;
    border-bottom: 1px solid var(--border);
    position: sticky;
    top: 0;
  }

  tbody tr {
    border-bottom: 1px solid rgba(42,42,61,0.5);
    transition: background 0.15s;
  }
  tbody tr:hover { background: rgba(108,99,255,0.06); }

  tbody td {
    padding: 10px 14px;
    color: var(--text);
    white-space: nowrap;
    max-width: 200px;
    overflow: hidden;
    text-overflow: ellipsis;
  }

  .table-container {
    max-height: 360px;
    overflow-y: auto;
    border-radius: 8px;
    border: 1px solid var(--border);
  }

  .table-container::-webkit-scrollbar { width: 6px; height: 6px; }
  .table-container::-webkit-scrollbar-track { background: var(--surface2); }
  .table-container::-webkit-scrollbar-thumb { background: var(--border); border-radius: 3px; }

  /* EXPORT BUTTONS */
  .export-bar {
    display: flex;
    gap: 12px;
    flex-wrap: wrap;
    margin-bottom: 36px;
    padding: 20px;
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    align-items: center;
  }

  .export-label {
    font-family: 'Syne', sans-serif;
    font-size: 12px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    color: var(--text2);
    margin-right: 4px;
  }

  .btn {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 10px 20px;
    border-radius: 8px;
    font-size: 13px;
    font-weight: 500;
    cursor: pointer;
    border: none;
    transition: all 0.2s;
    font-family: 'DM Sans', sans-serif;
    position: relative;
    overflow: hidden;
  }

  .btn::after {
    content: '';
    position: absolute;
    inset: 0;
    background: white;
    opacity: 0;
    transition: opacity 0.2s;
  }
  .btn:hover::after { opacity: 0.07; }
  .btn:active { transform: scale(0.97); }

  .btn-pdf {
    background: linear-gradient(135deg, #e53e3e, #c53030);
    color: white;
  }

  .btn-excel {
    background: linear-gradient(135deg, #2f855a, #276749);
    color: white;
  }

  .btn-png {
    background: linear-gradient(135deg, #6c63ff, #5a52d5);
    color: white;
  }

  .btn:disabled {
    opacity: 0.4;
    cursor: not-allowed;
    pointer-events: none;
  }

  /* COLUMN CONFIG */
  .config-section {
    display: none;
    margin-bottom: 28px;
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 22px;
  }
  .config-section.visible { display: block; }

  .config-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
    gap: 12px;
    margin-top: 14px;
  }

  .config-item {
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 12px 14px;
    display: flex;
    align-items: center;
    gap: 10px;
    cursor: pointer;
    transition: all 0.2s;
    font-size: 13px;
  }
  .config-item:hover { border-color: var(--accent); }
  .config-item.selected { border-color: var(--accent); background: rgba(108,99,255,0.12); }

  .config-check {
    width: 18px; height: 18px;
    border-radius: 5px;
    border: 2px solid var(--border);
    background: transparent;
    display: flex; align-items: center; justify-content: center;
    flex-shrink: 0;
    transition: all 0.2s;
  }
  .config-item.selected .config-check {
    background: var(--accent);
    border-color: var(--accent);
  }
  .config-check::after {
    content: '✓';
    font-size: 11px;
    color: white;
    opacity: 0;
    transition: opacity 0.2s;
  }
  .config-item.selected .config-check::after { opacity: 1; }

  .config-col-name {
    font-size: 12px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .apply-btn {
    margin-top: 16px;
    background: linear-gradient(135deg, var(--accent), #8b7ff5);
    color: white;
    border: none;
    padding: 10px 22px;
    border-radius: 8px;
    font-size: 13px;
    font-weight: 600;
    cursor: pointer;
    transition: all 0.2s;
    font-family: 'DM Sans', sans-serif;
  }
  .apply-btn:hover { opacity: 0.9; }

  /* CHART TYPE SELECTOR */
  .chart-type-row {
    display: flex;
    gap: 6px;
    margin-bottom: 14px;
  }
  .ct-btn {
    padding: 5px 12px;
    border-radius: 6px;
    font-size: 12px;
    cursor: pointer;
    border: 1px solid var(--border);
    background: var(--surface2);
    color: var(--text2);
    transition: all 0.2s;
  }
  .ct-btn.active {
    background: var(--accent);
    border-color: var(--accent);
    color: white;
  }

  /* EMPTY STATE */
  .empty-state {
    text-align: center;
    padding: 80px 20px;
    color: var(--text2);
  }
  .empty-icon { font-size: 56px; margin-bottom: 16px; opacity: 0.4; }
  .empty-title { font-family: 'Syne', sans-serif; font-size: 18px; font-weight: 700; margin-bottom: 8px; color: var(--text); opacity: 0.5; }
  .empty-sub { font-size: 14px; }

  /* LOADING */
  .loading-bar {
    height: 3px;
    background: linear-gradient(90deg, var(--accent), var(--accent2), var(--accent3));
    position: fixed;
    top: 0; left: 0;
    border-radius: 2px;
    transition: width 0.3s;
    z-index: 999;
  }

  /* TOAST */
  .toast {
    position: fixed;
    bottom: 24px; right: 24px;
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 14px 20px;
    font-size: 13px;
    display: flex;
    align-items: center;
    gap: 10px;
    transform: translateY(80px);
    opacity: 0;
    transition: all 0.3s;
    z-index: 9999;
    max-width: 320px;
    box-shadow: 0 8px 32px rgba(0,0,0,0.4);
  }
  .toast.show { transform: translateY(0); opacity: 1; }
  .toast-icon { font-size: 18px; }

  /* PAGINATION */
  .pagination {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-top: 14px;
    font-size: 12px;
    color: var(--text2);
  }
  .page-btns { display: flex; gap: 6px; }
  .page-btn {
    width: 32px; height: 32px;
    border-radius: 6px;
    border: 1px solid var(--border);
    background: var(--surface2);
    color: var(--text);
    cursor: pointer;
    display: flex; align-items: center; justify-content: center;
    font-size: 13px;
    transition: all 0.2s;
  }
  .page-btn.active { background: var(--accent); border-color: var(--accent); color: white; }
  .page-btn:hover:not(.active) { border-color: var(--accent); }
  .page-btn:disabled { opacity: 0.3; cursor: not-allowed; }

  /* RESPONSIVE */
  @media (max-width: 768px) {
    .dashboard-grid { grid-template-columns: 1fr; }
    .kpi-grid { grid-template-columns: repeat(2, 1fr); }
    header { flex-direction: column; gap: 12px; align-items: flex-start; }
  }

  /* HIDDEN CAPTURE AREA */
  #captureArea { background: #0a0a0f; }
</style>
</head>
<body>
<div class="loading-bar" id="loadingBar" style="width:0"></div>

<div class="app" id="captureArea">
  <!-- HEADER -->
  <header>
    <div class="logo">
      <div class="logo-icon">📊</div>
      <div class="logo-text">Data<span>Dash</span></div>
    </div>
    <div class="header-badge" id="fileStatus">Sin archivo cargado</div>
  </header>

  <!-- UPLOAD -->
  <div class="upload-section">
    <div class="section-label">01 · Fuente de datos</div>
    <div class="upload-zone" id="uploadZone">
      <span class="upload-icon">📁</span>
      <div class="upload-title">Arrastra tu archivo Excel aquí</div>
      <div class="upload-sub">Soporta .xlsx, .xls, .csv · Máx 50MB</div>
      <label class="upload-btn" for="fileInput">Seleccionar archivo</label>
      <input type="file" id="fileInput" accept=".xlsx,.xls,.csv">
    </div>
    <div class="file-info" id="fileInfo">
      <div class="file-dot"></div>
      <span class="file-name" id="fileName">—</span>
      <span class="file-meta" id="fileMeta">—</span>
    </div>
    <div class="sheet-selector" id="sheetSelector"></div>
  </div>

  <!-- CONFIG -->
  <div class="config-section" id="configSection">
    <div class="section-label">02 · Configurar columnas para gráficas</div>
    <p style="font-size:13px; color:var(--text2); margin-top:6px;">Selecciona las columnas que deseas visualizar en las gráficas:</p>
    <div class="config-grid" id="configGrid"></div>
    <button class="apply-btn" onclick="applyConfig()">Aplicar configuración →</button>
  </div>

  <!-- EXPORT -->
  <div class="export-bar" id="exportBar" style="display:none">
    <span class="export-label">Exportar</span>
    <button class="btn btn-pdf" onclick="exportPDF()">
      <span>📄</span> PDF
    </button>
    <button class="btn btn-png" onclick="exportPNG()">
      <span>🖼️</span> PNG
    </button>
    <button class="btn btn-excel" onclick="exportExcel()">
      <span>📊</span> Excel
    </button>
  </div>

  <!-- KPI CARDS -->
  <div id="kpiSection"></div>

  <!-- DASHBOARD -->
  <div id="dashboardSection">
    <div class="empty-state">
      <div class="empty-icon">📈</div>
      <div class="empty-title">Esperando datos</div>
      <div class="empty-sub">Carga un archivo Excel para generar el dashboard</div>
    </div>
  </div>
</div>

<!-- TOAST -->
<div class="toast" id="toast">
  <span class="toast-icon" id="toastIcon">✅</span>
  <span id="toastMsg">Mensaje</span>
</div>

<script>
// ─── STATE ─────────────────────────────────────
let workbook = null;
let currentSheet = null;
let allData = [];
let headers = [];
let selectedCols = [];
let charts = {};
let currentPage = 1;
const ROWS_PER_PAGE = 20;
let chartType1 = 'bar';
let chartType2 = 'line';

// ─── UPLOAD ZONE ───────────────────────────────
const uploadZone = document.getElementById('uploadZone');
const fileInput = document.getElementById('fileInput');

uploadZone.addEventListener('dragover', e => { e.preventDefault(); uploadZone.classList.add('drag-over'); });
uploadZone.addEventListener('dragleave', () => uploadZone.classList.remove('drag-over'));
uploadZone.addEventListener('drop', e => {
  e.preventDefault();
  uploadZone.classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) processFile(file);
});

fileInput.addEventListener('change', e => {
  if (e.target.files[0]) processFile(e.target.files[0]);
});

uploadZone.addEventListener('click', e => {
  if (e.target !== fileInput && !e.target.classList.contains('upload-btn')) {
    fileInput.click();
  }
});

function processFile(file) {
  if (!file) return;
  setLoading(true);
  const ext = file.name.split('.').pop().toLowerCase();

  document.getElementById('fileName').textContent = file.name;
  document.getElementById('fileMeta').textContent = formatBytes(file.size);
  document.getElementById('fileInfo').classList.add('visible');
  document.getElementById('fileStatus').textContent = file.name;

  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = new Uint8Array(e.target.result);
      workbook = XLSX.read(data, { type: 'array', cellDates: true });
      renderSheetTabs();
      loadSheet(workbook.SheetNames[0]);
      setLoading(false);
      showToast('✅', `Archivo cargado: ${workbook.SheetNames.length} hoja(s) encontrada(s)`);
    } catch(err) {
      setLoading(false);
      showToast('❌', 'Error al leer el archivo. Verifica el formato.');
    }
  };
  reader.readAsArrayBuffer(file);
}

function renderSheetTabs() {
  const sel = document.getElementById('sheetSelector');
  sel.innerHTML = '';
  sel.classList.add('visible');
  workbook.SheetNames.forEach((name, i) => {
    const tab = document.createElement('button');
    tab.className = 'sheet-tab' + (i === 0 ? ' active' : '');
    tab.textContent = name;
    tab.onclick = () => {
      document.querySelectorAll('.sheet-tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      loadSheet(name);
    };
    sel.appendChild(tab);
  });
}

function loadSheet(sheetName) {
  currentSheet = sheetName;
  const ws = workbook.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  if (raw.length < 2) {
    showToast('⚠️', 'La hoja no tiene suficientes datos');
    return;
  }

  headers = raw[0].map(h => String(h).trim() || `Col${Math.random().toString(36).substr(2,4)}`);
  allData = raw.slice(1).filter(row => row.some(c => c !== ''));
  currentPage = 1;

  renderConfigSection();
  renderKPIs();
  renderDashboard();
  document.getElementById('exportBar').style.display = 'flex';
}

// ─── CONFIG ─────────────────────────────────────
function renderConfigSection() {
  const grid = document.getElementById('configGrid');
  grid.innerHTML = '';
  document.getElementById('configSection').classList.add('visible');

  // Auto-detect numeric columns
  const numericCols = headers.filter((h, i) => {
    const vals = allData.map(r => r[i]).filter(v => v !== '');
    return vals.length > 0 && vals.filter(v => !isNaN(parseFloat(v))).length / vals.length > 0.6;
  });

  selectedCols = numericCols.slice(0, 4);

  headers.forEach((h, i) => {
    const item = document.createElement('div');
    const isNum = !isNaN(parseFloat(allData[0]?.[i]));
    item.className = 'config-item' + (selectedCols.includes(h) ? ' selected' : '');
    item.dataset.col = h;
    item.dataset.idx = i;
    item.innerHTML = `<div class="config-check"></div><span class="config-col-name">${h}</span><span style="font-size:10px;color:var(--text2);margin-left:auto">${isNum?'#':'Aa'}</span>`;
    item.onclick = () => {
      item.classList.toggle('selected');
      if (item.classList.contains('selected')) {
        if (!selectedCols.includes(h)) selectedCols.push(h);
      } else {
        selectedCols = selectedCols.filter(c => c !== h);
      }
    };
    grid.appendChild(item);
  });
}

function applyConfig() {
  renderDashboard();
  showToast('✅', 'Configuración aplicada');
}

// ─── KPIs ───────────────────────────────────────
function renderKPIs() {
  const numCols = headers.filter((h, i) => {
    const vals = allData.map(r => r[i]).filter(v => v !== '' && !isNaN(parseFloat(v)));
    return vals.length > allData.length * 0.5;
  }).slice(0, 4);

  const section = document.getElementById('kpiSection');
  if (!numCols.length) { section.innerHTML = ''; return; }

  let html = `<div class="section-label" style="margin-bottom:14px">03 · Indicadores clave</div><div class="kpi-grid">`;

  numCols.forEach(col => {
    const idx = headers.indexOf(col);
    const vals = allData.map(r => parseFloat(r[idx])).filter(v => !isNaN(v));
    const sum = vals.reduce((a,b)=>a+b,0);
    const avg = sum / vals.length;
    const max = Math.max(...vals);

    html += `
      <div class="kpi-card">
        <div class="kpi-label">${col}</div>
        <div class="kpi-value">${formatNum(sum)}</div>
        <div class="kpi-sub">Promedio: ${formatNum(avg)} · Máx: ${formatNum(max)}</div>
      </div>`;
  });

  html += `
    <div class="kpi-card">
      <div class="kpi-label">Total registros</div>
      <div class="kpi-value">${allData.length.toLocaleString()}</div>
      <div class="kpi-sub">${headers.length} columnas · Hoja: ${currentSheet}</div>
    </div>
  </div>`;

  section.innerHTML = html;
}

// ─── DASHBOARD ──────────────────────────────────
function renderDashboard() {
  destroyCharts();
  const section = document.getElementById('dashboardSection');
  const numCols = selectedCols.filter(c => {
    const i = headers.indexOf(c);
    return allData.filter(r => !isNaN(parseFloat(r[i]))).length > 0;
  });

  if (numCols.length === 0) {
    section.innerHTML = `<div class="empty-state">
      <div class="empty-icon">⚙️</div>
      <div class="empty-title">Selecciona columnas numéricas</div>
      <div class="empty-sub">Elige columnas numéricas en la configuración de arriba</div>
    </div>`;
    return;
  }

  const labelCol = headers.find(h => !selectedCols.includes(h)) || headers[0];
  const labelIdx = headers.indexOf(labelCol);
  const sampleSize = Math.min(20, allData.length);
  const labels = allData.slice(0, sampleSize).map(r => String(r[labelIdx] || '').slice(0, 20));

  section.innerHTML = `
    <div class="section-label" style="margin-bottom:16px">04 · Visualizaciones</div>
    <div class="dashboard-grid">
      <div class="panel">
        <div class="panel-header">
          <div class="panel-title">📊 Distribución por columna</div>
          <div class="panel-badge" id="badge1">${numCols[0] || ''}</div>
        </div>
        <div class="chart-type-row" id="ct1"></div>
        <div class="chart-wrap"><canvas id="chart1"></canvas></div>
      </div>

      <div class="panel">
        <div class="panel-header">
          <div class="panel-title">📈 Tendencia general</div>
          <div class="panel-badge">${numCols.length} series</div>
        </div>
        <div class="chart-type-row" id="ct2"></div>
        <div class="chart-wrap"><canvas id="chart2"></canvas></div>
      </div>

      ${numCols.length >= 2 ? `
      <div class="panel">
        <div class="panel-header">
          <div class="panel-title">🍩 Composición</div>
          <div class="panel-badge">Pie</div>
        </div>
        <div class="chart-wrap" style="height:240px"><canvas id="chart3"></canvas></div>
      </div>

      <div class="panel">
        <div class="panel-header">
          <div class="panel-title">📉 Comparativo</div>
          <div class="panel-badge">Radar</div>
        </div>
        <div class="chart-wrap" style="height:240px"><canvas id="chart4"></canvas></div>
      </div>
      ` : ''}

      <div class="panel full-width">
        <div class="panel-header">
          <div class="panel-title">📋 Datos completos</div>
          <div class="panel-badge" id="tableCount">${allData.length} filas</div>
        </div>
        <div class="table-container">
          <div class="table-scroll"><div id="tableWrapper"></div></div>
        </div>
        <div class="pagination" id="pagination"></div>
      </div>
    </div>`;

  // Chart type selectors
  renderChartTypeBtns('ct1', ['bar','horizontalBar','area'], chartType1, t => { chartType1 = t; buildChart1(numCols, labels); });
  renderChartTypeBtns('ct2', ['line','bar','scatter'], chartType2, t => { chartType2 = t; buildChart2(numCols, labels); });

  // Build charts
  buildChart1(numCols, labels);
  buildChart2(numCols, labels);
  if (numCols.length >= 2) {
    buildPieChart(numCols, labels);
    buildRadarChart(numCols);
  }

  renderTable();
}

function renderChartTypeBtns(containerId, types, active, onChange) {
  const c = document.getElementById(containerId);
  const labels = { bar: 'Barra', horizontalBar: 'H-Barra', area: 'Área', line: 'Línea', scatter: 'Dispersión' };
  types.forEach(t => {
    const btn = document.createElement('button');
    btn.className = 'ct-btn' + (t === active ? ' active' : '');
    btn.textContent = labels[t] || t;
    btn.onclick = () => {
      c.querySelectorAll('.ct-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      onChange(t);
    };
    c.appendChild(btn);
  });
}

// COLORS
const COLORS = [
  'rgba(108,99,255,0.85)', 'rgba(255,101,132,0.85)',
  'rgba(67,233,123,0.85)', 'rgba(247,151,30,0.85)',
  'rgba(56,189,248,0.85)', 'rgba(196,131,255,0.85)'
];

const BORDER_COLORS = [
  'rgb(108,99,255)', 'rgb(255,101,132)',
  'rgb(67,233,123)', 'rgb(247,151,30)',
  'rgb(56,189,248)', 'rgb(196,131,255)'
];

const CHART_DEFAULTS = {
  responsive: true,
  maintainAspectRatio: false,
  plugins: {
    legend: { labels: { color: '#8888aa', font: { family: 'DM Sans' }, boxWidth: 12 } },
    tooltip: { backgroundColor: '#1c1c28', borderColor: '#2a2a3d', borderWidth: 1, titleColor: '#e8e8f0', bodyColor: '#8888aa', padding: 10 }
  },
  scales: {
    x: { ticks: { color: '#8888aa', font: { size: 11 } }, grid: { color: 'rgba(42,42,61,0.5)' }, border: { color: '#2a2a3d' } },
    y: { ticks: { color: '#8888aa', font: { size: 11 } }, grid: { color: 'rgba(42,42,61,0.5)' }, border: { color: '#2a2a3d' } }
  }
};

function buildChart1(numCols, labels) {
  if (charts.c1) charts.c1.destroy();
  const col = numCols[0];
  const idx = headers.indexOf(col);
  const vals = allData.slice(0, labels.length).map(r => parseFloat(r[idx]) || 0);
  const ctx = document.getElementById('chart1');
  if (!ctx) return;

  let type = chartType1 === 'horizontalBar' ? 'bar' : (chartType1 === 'area' ? 'line' : chartType1);
  const isHoriz = chartType1 === 'horizontalBar';

  charts.c1 = new Chart(ctx, {
    type,
    data: {
      labels,
      datasets: [{
        label: col,
        data: vals,
        backgroundColor: chartType1 === 'area' ? 'rgba(108,99,255,0.2)' : COLORS[0],
        borderColor: BORDER_COLORS[0],
        borderWidth: 2,
        fill: chartType1 === 'area',
        tension: 0.4,
        borderRadius: type === 'bar' && !isHoriz ? 6 : 0
      }]
    },
    options: {
      ...CHART_DEFAULTS,
      indexAxis: isHoriz ? 'y' : 'x',
    }
  });
  document.getElementById('badge1').textContent = col;
}

function buildChart2(numCols, labels) {
  if (charts.c2) charts.c2.destroy();
  const ctx = document.getElementById('chart2');
  if (!ctx) return;

  const datasets = numCols.map((col, i) => {
    const idx = headers.indexOf(col);
    const vals = allData.slice(0, labels.length).map(r => parseFloat(r[idx]) || 0);
    return {
      label: col,
      data: chartType2 === 'scatter' ? vals.map((v, j) => ({ x: j, y: v })) : vals,
      backgroundColor: COLORS[i % COLORS.length],
      borderColor: BORDER_COLORS[i % BORDER_COLORS.length],
      borderWidth: 2,
      tension: 0.4,
      borderRadius: chartType2 === 'bar' ? 4 : 0,
      fill: false,
      pointRadius: chartType2 === 'scatter' ? 5 : 3
    };
  });

  charts.c2 = new Chart(ctx, {
    type: chartType2 === 'scatter' ? 'scatter' : chartType2,
    data: { labels: chartType2 !== 'scatter' ? labels : undefined, datasets },
    options: CHART_DEFAULTS
  });
}

function buildPieChart(numCols, labels) {
  if (charts.c3) charts.c3.destroy();
  const ctx = document.getElementById('chart3');
  if (!ctx) return;

  const totals = numCols.map(col => {
    const idx = headers.indexOf(col);
    return allData.reduce((s, r) => s + (parseFloat(r[idx]) || 0), 0);
  });

  charts.c3 = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: numCols,
      datasets: [{ data: totals, backgroundColor: COLORS, borderColor: '#13131a', borderWidth: 3 }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'right', labels: { color: '#8888aa', font: { family: 'DM Sans' }, boxWidth: 12 } },
        tooltip: CHART_DEFAULTS.plugins.tooltip
      }
    }
  });
}

function buildRadarChart(numCols) {
  if (charts.c4) charts.c4.destroy();
  const ctx = document.getElementById('chart4');
  if (!ctx) return;

  const sample = allData.slice(0, 5);
  const labelIdx = headers.findIndex(h => !selectedCols.includes(h));

  const datasets = sample.map((row, i) => {
    const vals = numCols.map(col => {
      const idx = headers.indexOf(col);
      return parseFloat(row[idx]) || 0;
    });
    const nm = labelIdx >= 0 ? String(row[labelIdx]).slice(0, 15) : `Fila ${i+1}`;
    return {
      label: nm,
      data: vals,
      backgroundColor: COLORS[i % COLORS.length].replace('0.85', '0.2'),
      borderColor: BORDER_COLORS[i % BORDER_COLORS.length],
      borderWidth: 2,
      pointBackgroundColor: BORDER_COLORS[i % BORDER_COLORS.length]
    };
  });

  charts.c4 = new Chart(ctx, {
    type: 'radar',
    data: { labels: numCols, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { labels: { color: '#8888aa', font: { family: 'DM Sans' }, boxWidth: 12 } }, tooltip: CHART_DEFAULTS.plugins.tooltip },
      scales: { r: { ticks: { color: '#8888aa', backdropColor: 'transparent' }, grid: { color: 'rgba(42,42,61,0.6)' }, angleLines: { color: 'rgba(42,42,61,0.8)' }, pointLabels: { color: '#8888aa', font: { size: 11 } } } }
    }
  });
}

// ─── TABLE ──────────────────────────────────────
function renderTable() {
  const wrapper = document.getElementById('tableWrapper');
  const pagination = document.getElementById('pagination');
  if (!wrapper) return;

  const totalPages = Math.ceil(allData.length / ROWS_PER_PAGE);
  const start = (currentPage - 1) * ROWS_PER_PAGE;
  const pageData = allData.slice(start, start + ROWS_PER_PAGE);

  let html = '<table><thead><tr>';
  headers.forEach(h => { html += `<th>${h}</th>`; });
  html += '</tr></thead><tbody>';
  pageData.forEach(row => {
    html += '<tr>';
    headers.forEach((_, i) => { html += `<td>${row[i] ?? ''}</td>`; });
    html += '</tr>';
  });
  html += '</tbody></table>';
  wrapper.innerHTML = html;

  // Pagination
  let pag = `<span>${allData.length} registros · Página ${currentPage} de ${totalPages}</span>`;
  pag += '<div class="page-btns">';
  pag += `<button class="page-btn" onclick="goPage(${currentPage-1})" ${currentPage===1?'disabled':''}>‹</button>`;
  const pages = getPageRange(currentPage, totalPages);
  pages.forEach(p => {
    if (p === '...') { pag += `<span style="padding:0 4px;color:var(--text2)">…</span>`; }
    else { pag += `<button class="page-btn${p===currentPage?' active':''}" onclick="goPage(${p})">${p}</button>`; }
  });
  pag += `<button class="page-btn" onclick="goPage(${currentPage+1})" ${currentPage===totalPages?'disabled':''}>›</button>`;
  pag += '</div>';
  pagination.innerHTML = pag;
}

function goPage(p) {
  const total = Math.ceil(allData.length / ROWS_PER_PAGE);
  if (p < 1 || p > total) return;
  currentPage = p;
  renderTable();
}

function getPageRange(current, total) {
  if (total <= 7) return Array.from({ length: total }, (_, i) => i+1);
  if (current <= 4) return [1,2,3,4,5,'...',total];
  if (current >= total - 3) return [1,'...',total-4,total-3,total-2,total-1,total];
  return [1,'...',current-1,current,current+1,'...',total];
}

// ─── EXPORT ─────────────────────────────────────
async function exportPDF() {
  showToast('⏳', 'Generando PDF...');
  const { jsPDF } = window.jspdf;

  try {
    const el = document.getElementById('captureArea');
    const canvas = await html2canvas(el, { scale: 1.5, backgroundColor: '#0a0a0f', useCORS: true, logging: false });
    const imgData = canvas.toDataURL('image/jpeg', 0.9);

    const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    const pageW = pdf.internal.pageSize.getWidth();
    const pageH = pdf.internal.pageSize.getHeight();
    const imgW = canvas.width;
    const imgH = canvas.height;
    const ratio = Math.min(pageW / imgW, pageH / imgH);
    const w = imgW * ratio;
    const h = imgH * ratio;
    const x = (pageW - w) / 2;
    const y = (pageH - h) / 2;

    pdf.addImage(imgData, 'JPEG', x, y, w, h);
    pdf.save('dashboard_informe.pdf');
    showToast('✅', 'PDF descargado correctamente');
  } catch (e) {
    showToast('❌', 'Error al generar PDF');
  }
}

async function exportPNG() {
  showToast('⏳', 'Generando imagen...');
  try {
    const el = document.getElementById('captureArea');
    const canvas = await html2canvas(el, { scale: 2, backgroundColor: '#0a0a0f', useCORS: true, logging: false });
    const link = document.createElement('a');
    link.download = 'dashboard_informe.png';
    link.href = canvas.toDataURL('image/png');
    link.click();
    showToast('✅', 'PNG descargado correctamente');
  } catch(e) {
    showToast('❌', 'Error al generar PNG');
  }
}

function exportExcel() {
  if (!allData.length) return;
  showToast('⏳', 'Generando Excel...');

  try {
    const wb = XLSX.utils.book_new();

    // Sheet 1: Raw data
    const wsData = [headers, ...allData];
    const ws1 = XLSX.utils.aoa_to_sheet(wsData);

    // Style header row
    const headerRange = XLSX.utils.decode_range(ws1['!ref']);
    for (let c = headerRange.s.c; c <= headerRange.e.c; c++) {
      const cellAddr = XLSX.utils.encode_cell({ r: 0, c });
      if (!ws1[cellAddr]) continue;
      ws1[cellAddr].s = {
        font: { bold: true, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: '6C63FF' } },
        alignment: { horizontal: 'center' }
      };
    }

    // Col widths
    ws1['!cols'] = headers.map(h => ({ wch: Math.max(h.length + 4, 12) }));
    XLSX.utils.book_append_sheet(wb, ws1, 'Datos');

    // Sheet 2: Summary / KPIs
    const numericHeaders = headers.filter((h, i) => {
      const vals = allData.map(r => r[i]).filter(v => v !== '' && !isNaN(parseFloat(v)));
      return vals.length > allData.length * 0.5;
    });

    const summaryData = [
      ['Resumen Ejecutivo', '', '', '', ''],
      ['Generado:', new Date().toLocaleString('es-CO'), '', '', ''],
      ['Hoja de origen:', currentSheet, '', '', ''],
      ['Total registros:', allData.length, '', '', ''],
      ['Total columnas:', headers.length, '', '', ''],
      ['', '', '', '', ''],
      ['Columna', 'Suma', 'Promedio', 'Mínimo', 'Máximo', 'Conteo']
    ];

    numericHeaders.forEach(h => {
      const idx = headers.indexOf(h);
      const vals = allData.map(r => parseFloat(r[idx])).filter(v => !isNaN(v));
      if (!vals.length) return;
      const sum = vals.reduce((a, b) => a + b, 0);
      summaryData.push([h, sum, sum / vals.length, Math.min(...vals), Math.max(...vals), vals.length]);
    });

    const ws2 = XLSX.utils.aoa_to_sheet(summaryData);
    ws2['!cols'] = [{ wch: 20 }, { wch: 18 }, { wch: 18 }, { wch: 14 }, { wch: 14 }, { wch: 10 }];
    XLSX.utils.book_append_sheet(wb, ws2, 'Resumen');

    XLSX.writeFile(wb, 'dashboard_informe.xlsx');
    showToast('✅', 'Excel generado con 2 hojas: Datos y Resumen');
  } catch(e) {
    showToast('❌', 'Error al generar Excel');
  }
}

// ─── HELPERS ────────────────────────────────────
function destroyCharts() {
  Object.values(charts).forEach(c => { try { c.destroy(); } catch(e){} });
  charts = {};
}

function formatBytes(b) {
  if (b < 1024) return b + ' B';
  if (b < 1048576) return (b/1024).toFixed(1) + ' KB';
  return (b/1048576).toFixed(1) + ' MB';
}

function formatNum(n) {
  if (isNaN(n)) return '—';
  if (Math.abs(n) >= 1e9) return (n/1e9).toFixed(1) + 'B';
  if (Math.abs(n) >= 1e6) return (n/1e6).toFixed(1) + 'M';
  if (Math.abs(n) >= 1e3) return (n/1e3).toFixed(1) + 'K';
  return n % 1 === 0 ? n.toLocaleString('es-CO') : n.toFixed(2);
}

let toastTimer = null;
function showToast(icon, msg) {
  const toast = document.getElementById('toast');
  document.getElementById('toastIcon').textContent = icon;
  document.getElementById('toastMsg').textContent = msg;
  toast.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove('show'), 3500);
}

function setLoading(on) {
  const bar = document.getElementById('loadingBar');
  if (on) {
    bar.style.width = '70%';
    bar.style.transition = 'width 0.8s ease';
  } else {
    bar.style.width = '100%';
    setTimeout(() => { bar.style.width = '0'; bar.style.transition = 'none'; }, 400);
  }
}
</script>
</body>
</html>
