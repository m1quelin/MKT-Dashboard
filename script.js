const MONTH_NAMES = ['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO','JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'];
const MONTH_COLS = {
  JANEIRO:[2,3],FEVEREIRO:[5,6],MARÇO:[8,9],ABRIL:[11,12],
  MAIO:[14,15],JUNHO:[17,18],JULHO:[20,21],AGOSTO:[23,24],
  SETEMBRO:[26,27],OUTUBRO:[29,30],NOVEMBRO:[32,33],DEZEMBRO:[35,36]
};
const LS_KEY = 'dashboard_mkt_xlsx_b64';

let globalData = null;
let currentModalMonth = null;
let currentModalRows = [];

// ─── AUTO-LOAD FROM LOCALSTORAGE ───────────────────────────────────────────
window.addEventListener('DOMContentLoaded', () => {
  const saved = localStorage.getItem(LS_KEY);
  if (saved) {
    try {
      const wb = XLSX.read(saved, { type: 'base64', cellDates: true });
      parseAndRender(wb, true);
    } catch(e) {
      localStorage.removeItem(LS_KEY);
    }
  }
});

// ─── FILE INPUT ──────────────────────────────────────────────────────────────
document.getElementById('fileInput').addEventListener('change', e => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const b64 = ev.target.result.split(',')[1];
      localStorage.setItem(LS_KEY, b64);
      const wb = XLSX.read(b64, { type: 'base64', cellDates: true });
      parseAndRender(wb, false);
    } catch(err) {
      alert('Erro ao ler planilha: ' + err.message);
    }
  };
  reader.readAsDataURL(file);
});

// ─── PARSE ───────────────────────────────────────────────────────────────────
function parseAndRender(wb, fromCache) {
  const mktSheet = wb.Sheets[wb.SheetNames[0]];
  const mktRaw = XLSX.utils.sheet_to_json(mktSheet, { header: 1, defval: null });

  const totalRowIdx = mktRaw.findIndex(r => r[0] === 'TOTAL');
  const totalRow = totalRowIdx >= 0 ? mktRaw[totalRowIdx] : mktRaw[9] || [];
  const catRows = mktRaw.slice(2, totalRowIdx >= 0 ? totalRowIdx : 9).filter(r => r[0]);
  const orcadoAnual = toNum(totalRow[1]) || 2640000;

  const months = MONTH_NAMES.map(m => {
    const [oc, rc] = MONTH_COLS[m];
    const orcado = toNum(totalRow[oc]) || 0;
    const realizado = toNum(totalRow[rc]) || 0;
    const categories = catRows.map(r => ({
      name: r[0], realizado: toNum(r[rc]) || 0
    })).filter(c => c.realizado > 0);
    return { month: m, orcado, realizado, categories, hasData: realizado > 0 };
  });

  const lancSheet = wb.Sheets[wb.SheetNames[1]];
  const lancRaw = XLSX.utils.sheet_to_json(lancSheet, { header: 1, defval: null, cellDates: true });
  const lancamentos = lancRaw.slice(1).filter(r => r[1]).map(r => ({
    data: parseDate(r[1]),
    historico: r[2] || '',
    filial: r[3] || '',
    debito: toNum(r[4]) || 0,
    categoria: r[5] || ''
  }));

  // Detect file name hint from sheet name
  const sheetName = wb.SheetNames[0];
  document.getElementById('headerSub').textContent = fromCache
    ? 'Orçamento × Realizado  ·  dados salvos localmente'
    : 'Orçamento × Realizado';

  globalData = { months, lancamentos, orcadoAnual };
  renderDashboard();

  document.getElementById('pdfBtn').disabled = false;
  document.getElementById('savedBadge').style.display = 'flex';
  if (!fromCache) {
    setTimeout(() => document.getElementById('savedBadge').style.display = 'none', 4000);
  }
}

// ─── HELPERS ─────────────────────────────────────────────────────────────────
function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  if (typeof val === 'number') return XLSX.SSF.parse_date_code(val);
  return new Date(val);
}
function formatDate(d) {
  if (!d) return '';
  try { return (d instanceof Date ? d : new Date(d)).toLocaleDateString('pt-BR'); }
  catch { return ''; }
}
function toNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  const n = parseFloat(String(v).replace(',', '.'));
  return isNaN(n) ? 0 : n;
}
function fmt(n) { return n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL', minimumFractionDigits: 2 }); }
function fmtShort(n) {
  if (n >= 1e6) return 'R$ ' + (n/1e6).toFixed(2).replace('.',',') + 'M';
  if (n >= 1e3) return 'R$ ' + (n/1e3).toFixed(1).replace('.',',') + 'K';
  return fmt(n);
}
function tagClass(cat) {
  if (!cat) return 'tag-out';
  const c = cat.toLowerCase();
  if (c.includes('publicidade') || c.includes('propaganda')) return 'tag-pub';
  if (c.includes('brinde')) return 'tag-bri';
  if (c.includes('patroc')) return 'tag-pat';
  if (c.includes('doa')) return 'tag-doa';
  return 'tag-out';
}

// ─── RENDER ───────────────────────────────────────────────────────────────────
function renderDashboard() {
  const { months, lancamentos, orcadoAnual } = globalData;
  const realizados = months.filter(m => m.hasData);
  const totalRealizado = realizados.reduce((s, m) => s + m.realizado, 0);
  const totalOrcadoAte = realizados.reduce((s, m) => s + m.orcado, 0);
  const pctGeral = totalOrcadoAte > 0 ? (totalRealizado / totalOrcadoAte * 100) : 0;
  const saldo = totalOrcadoAte - totalRealizado;

  document.getElementById('main-content').innerHTML = `
    <div class="summary-bar">
      <div class="kpi-card">
        <div class="kpi-label">Orçado Anual</div>
        <div class="kpi-value">${fmtShort(orcadoAnual)}</div>
        <div class="kpi-sub">Total previsto 2026</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Realizado Acumulado</div>
        <div class="kpi-value ${pctGeral > 100 ? 'red' : 'green'}">${fmtShort(totalRealizado)}</div>
        <div class="kpi-sub">${realizados.length} ${realizados.length === 1 ? 'mês' : 'meses'} realizados</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">% Orçado (meses realizados)</div>
        <div class="kpi-value ${pctGeral > 110 ? 'red' : pctGeral > 100 ? 'yellow' : 'green'}">${pctGeral.toFixed(1)}%</div>
        <div class="kpi-sub">vs orçado dos meses</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Saldo (orçado − realizado)</div>
        <div class="kpi-value ${saldo < 0 ? 'red' : 'green'}">${fmtShort(Math.abs(saldo))}</div>
        <div class="kpi-sub">${saldo < 0 ? '▲ Acima do orçado' : '▼ Abaixo do orçado'}</div>
      </div>
    </div>

    <div class="section-title">Meses — clique para ver lançamentos de NF</div>
    <div class="months-grid">
      ${months.map(renderMonthCard).join('')}
    </div>

    <div class="section-title">Resumo por Categoria (acumulado)</div>
    ${renderCategoryTable()}
  `;

  document.querySelectorAll('.month-card[data-month]').forEach(card => {
    card.addEventListener('click', () => openModal(card.dataset.month));
  });
}

function renderMonthCard(m) {
  const pct = m.orcado > 0 ? (m.realizado / m.orcado * 100) : 0;
  const barColor = pct > 110 ? 'var(--red)' : pct > 100 ? 'var(--yellow)' : 'var(--green)';
  const barW = Math.min(pct, 100).toFixed(1);
  if (!m.hasData) return `<div class="month-card no-data">
    <div class="month-name">${m.month.toLowerCase()}</div>
    <div class="month-val" style="color:var(--muted)">${fmtShort(m.orcado)}</div>
    <div class="month-val-label">orçado</div>
    <div class="badge-future">Sem lançamentos</div>
  </div>`;
  return `<div class="month-card" data-month="${m.month}">
    <div class="month-name">${m.month.toLowerCase()}</div>
    <div class="month-val" style="color:${barColor}">${fmtShort(m.realizado)}</div>
    <div class="month-val-label">realizado de ${fmtShort(m.orcado)}</div>
    <div class="progress-bar-wrap"><div class="progress-bar" style="width:${barW}%;background:${barColor}"></div></div>
    <div class="month-pct" style="color:${barColor}">${pct.toFixed(1)}%</div>
  </div>`;
}

function renderCategoryTable() {
  const catMap = {};
  globalData.months.forEach(m => m.categories.forEach(c => {
    catMap[c.name] = (catMap[c.name] || 0) + c.realizado;
  }));
  const cats = Object.entries(catMap).sort((a,b) => b[1]-a[1]);
  const total = cats.reduce((s,c) => s+c[1], 0);
  const rows = cats.map(([name, val]) => `<tr>
    <td><span class="tag ${tagClass(name)}">${name}</span></td>
    <td class="td-right">${fmt(val)}</td>
    <td class="td-right">${total > 0 ? (val/total*100).toFixed(1) : '0.0'}%</td>
  </tr>`).join('');
  return `<div class="table-wrap">
    <div class="table-header">
      <span>Distribuição de Gastos Realizados</span>
      <span style="font-size:12px;color:var(--muted)">Total: ${fmt(total)}</span>
    </div>
    <table>
      <thead><tr>
        <th>Categoria</th>
        <th style="text-align:right">Realizado (R$)</th>
        <th style="text-align:right">Participação</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>
  </div>`;
}

// ─── MODAL ───────────────────────────────────────────────────────────────────
function openModal(monthName) {
  const MONTH_NUM = MONTH_NAMES.indexOf(monthName) + 1;
  currentModalMonth = monthName;
  const filtered = globalData.lancamentos.filter(l => {
    if (!l.data) return false;
    const dt = l.data instanceof Date ? l.data : new Date(l.data);
    return dt.getMonth() + 1 === MONTH_NUM;
  });
  currentModalRows = filtered;
  const total = filtered.reduce((s,l) => s+l.debito, 0);
  document.getElementById('modalTitle').textContent = `Lançamentos — ${monthName.charAt(0)+monthName.slice(1).toLowerCase()}`;
  document.getElementById('modalSub').textContent = `${filtered.length} NFs · Total: ${fmt(total)}`;
  renderModalTable(filtered);
  document.getElementById('modalOverlay').classList.add('open');
  document.getElementById('modalSearch').value = '';
  document.getElementById('modalSearch').oninput = e => {
    const q = e.target.value.toLowerCase();
    renderModalTable(q ? filtered.filter(l =>
      l.historico.toLowerCase().includes(q) ||
      l.categoria.toLowerCase().includes(q) ||
      String(l.filial).toLowerCase().includes(q)
    ) : filtered);
  };
}

async function captureToClipboard(label) {
  const loading = document.getElementById('pdfLoading');
  const loadingMsg = document.getElementById('pdfLoadingMsg');

  if (!loading || !loadingMsg) return;

  try {
    loadingMsg.textContent = `Capturando ${label || 'tela'}...`;
    loading.classList.add('show');

    // 👇 Esconde overlay ANTES do print
    loading.classList.remove('show');
    await new Promise(r => setTimeout(r, 50));

    const target = document.getElementById('main-content');
    if (!target) throw new Error('main-content não encontrado');

    const canvas = await html2canvas(target, {
      scale: window.devicePixelRatio || 1,
      useCORS: true,
      backgroundColor: '#0f1117'
    });

    // 👇 Reativa overlay depois da captura
    loading.classList.add('show');
    loadingMsg.textContent = 'Copiando imagem...';

    const blob = await new Promise((resolve, reject) => {
      canvas.toBlob(b => b ? resolve(b) : reject(new Error('Falha ao gerar imagem')), 'image/png');
    });

    await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);

    loadingMsg.textContent = 'Imagem copiada! Agora é só colar (Ctrl+V)';
    await new Promise(r => setTimeout(r, 2000));

  } catch (err) {
    console.error(err);
    alert('Erro ao capturar: ' + err.message);
  } finally {
    loading.classList.remove('show');
  }
}

// ─── BOTÃO PRINCIPAL ───────────────────────────────────────────
document.getElementById('pdfBtn').addEventListener('click', async () => {
  await captureToClipboard('Dashboard MKT');
});

// ─── BOTÃO DO MODAL ────────────────────────────────────────────
document.getElementById('modalPdfBtn').addEventListener('click', async () => {
  if (!currentModalMonth) return;
  await captureToClipboard('Lançamentos — ' + currentModalMonth);
});