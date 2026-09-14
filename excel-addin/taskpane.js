/* ==========================================================================
   DocBrisk KPI Dashboards — task pane controller
   ========================================================================== */

/* ---------------------------------------------------------------- state */
const State = {
  headers: [],
  rows: [],
  profiles: [],
  suggestions: [],   // KPI objects from KpiEngine, plus any custom ones appended
  selectedIds: new Set(),
  templateId: null
};

function $(sel) { return document.querySelector(sel); }
function showStep(n) { $(`.db-step[data-step="${n}"]`).classList.remove('hidden'); }
function setStatus(id, msg, kind) {
  const el = $('#' + id);
  el.textContent = msg || '';
  el.className = 'db-status' + (kind ? ' ' + kind : '');
}

/* ------------------------------------------------------------ step 1: data */
function wireStep1() {
  $('#btnUseSelection').addEventListener('click', () => loadData(false));
  $('#btnUseSheet').addEventListener('click', () => loadData(true));
}

async function loadData(wholeSheet) {
  setStatus('dataStatus', 'Reading…');
  try {
    await Excel.run(async (context) => {
      const range = wholeSheet
        ? context.workbook.worksheets.getActiveWorksheet().getUsedRange()
        : context.workbook.getSelectedRange();
      range.load('values, rowCount, columnCount');
      await context.sync();

      if (range.rowCount < 2) {
        throw new Error('Select at least a header row and one data row.');
      }
      const values = range.values;
      State.headers = values[0].map((h, i) => String(h ?? '').trim() || `Column ${i + 1}`);
      State.rows = values.slice(1);
      State.profiles = KpiEngine.profileColumns(State.headers, State.rows);
      State.suggestions = KpiEngine.suggestKpis(State.profiles, State.rows);
      // Pre-check the strongest handful so there's something useful to
      // build immediately; everything else stays available to opt into.
      State.selectedIds = new Set(State.suggestions.slice(0, 6).map((k) => k.id));
    });

    setStatus('dataStatus', `Loaded ${State.rows.length} rows, ${State.headers.length} columns.`, 'ok');
    renderKpiList();
    showStep(2); showStep(3); showStep(4);
  } catch (err) {
    setStatus('dataStatus', err.message || String(err), 'err');
  }
}

/* ------------------------------------------------------------ step 2: KPIs */
function wireStep2() {
  $('#btnAddCustom').addEventListener('click', addCustomKpi);
}

function renderKpiList() {
  const list = $('#kpiList');
  list.innerHTML = '';
  State.suggestions.forEach((kpi) => {
    const row = document.createElement('label');
    row.className = 'db-kpi-item';
    const checked = State.selectedIds.has(kpi.id);
    const valueText = kpi.value !== undefined ? formatKpiValue(kpi) : '';
    row.innerHTML = `
      <input type="checkbox" ${checked ? 'checked' : ''} data-id="${kpi.id}">
      <div>
        <span class="db-kpi-name">${escapeHtml(kpi.label)}</span>
        ${valueText ? `<span class="db-kpi-value">${escapeHtml(valueText)}</span>` : ''}
        <div class="db-kpi-explain">${escapeHtml(kpi.explain)}</div>
      </div>`;
    row.querySelector('input').addEventListener('change', (e) => {
      if (e.target.checked) State.selectedIds.add(kpi.id);
      else State.selectedIds.delete(kpi.id);
    });
    list.appendChild(row);
  });
}

function formatKpiValue(kpi) {
  if (kpi.format === 'currency') return '₹' + KpiEngine.formatNumber(kpi.value);
  if (kpi.format === 'percent') return KpiEngine.formatNumber(kpi.value) + '%';
  return KpiEngine.formatNumber(kpi.value);
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function addCustomKpi() {
  const label = $('#customLabel').value.trim();
  const formula = $('#customFormula').value.trim();
  const format = $('#customFormat').value;
  if (!label) { setStatus('customStatus', 'Give the KPI a name first.', 'err'); return; }
  try {
    const kpi = KpiEngine.buildCustomKpi(label, formula, State.profiles, State.rows, format);
    State.suggestions.unshift(kpi);
    State.selectedIds.add(kpi.id);
    renderKpiList();
    setStatus('customStatus', `Added "${label}".`, 'ok');
    $('#customLabel').value = ''; $('#customFormula').value = '';
  } catch (err) {
    setStatus('customStatus', err.message || String(err), 'err');
  }
}

/* --------------------------------------------------------- step 3: template */
function wireStep3() {
  const list = $('#templateList');
  KpiTemplates.list().forEach((tpl) => {
    const item = document.createElement('div');
    item.className = 'db-template-item';
    item.innerHTML = `<div><div class="db-template-name">${escapeHtml(tpl.name)}</div>
      <div class="db-template-desc">${escapeHtml(tpl.description)}</div></div>`;
    item.addEventListener('click', () => {
      State.templateId = tpl.id;
      list.querySelectorAll('.db-template-item').forEach((n) => n.classList.toggle('on', n === item));
    });
    list.appendChild(item);
  });
  // Default to the first template so Step 4 works even if the user never
  // opens Step 3 — they can still come back and change it before building.
  if (KpiTemplates.list().length) {
    State.templateId = KpiTemplates.list()[0].id;
    list.firstChild.classList.add('on');
  }
}

function currentFilledSlots() {
  const template = KpiTemplates.get(State.templateId);
  const chosen = State.suggestions.filter((k) => State.selectedIds.has(k.id));
  if (!chosen.length) throw new Error('Check at least one KPI in step 2 first.');
  return { template, filled: KpiTemplates.autoFill(template, chosen) };
}

/* ---------------------------------------------------------- step 4: build */
function wireStep4() {
  $('#btnBuildExcel').addEventListener('click', buildInExcel);
  $('#btnExportHtml').addEventListener('click', exportHtml);
  $('#btnDownloadCsv').addEventListener('click', downloadCsv);
  $('#btnLookerLink').addEventListener('click', openLooker);
}

async function buildInExcel() {
  setStatus('buildStatus', 'Building…');
  try {
    const { filled } = currentFilledSlots();
    await Excel.run((context) => DashboardBuilder.build(context, filled));
    setStatus('buildStatus', `Added a "${DashboardBuilder.SHEET_NAME}" sheet to this workbook.`, 'ok');
  } catch (err) {
    setStatus('buildStatus', err.message || String(err), 'err');
  }
}

function exportHtml() {
  try {
    const { filled } = currentFilledSlots();
    const html = DashboardExport.buildHtml('KPI Dashboard', filled);
    const blob = new Blob([html], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = 'dashboard.html';
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
    setStatus('buildStatus', 'Downloaded dashboard.html.', 'ok');
  } catch (err) {
    setStatus('buildStatus', err.message || String(err), 'err');
  }
}

function downloadCsv() {
  const csv = LookerStudio.buildCsv(State.headers, State.rows);
  LookerStudio.downloadCsv('dashboard-data.csv', csv);
  setStatus('lookerStatus', 'Downloaded — import this into a new Google Sheet.', 'ok');
}

function openLooker() {
  const template = KpiTemplates.get(State.templateId);
  if (!template.lookerReportId) {
    setStatus('lookerStatus',
      `No Looker Studio template is set up yet for "${template.name}". See the README's "Looker Studio setup" section — it's a one-time step.`, 'err');
    return;
  }
  const parsed = LookerStudio.parseSheetUrl($('#sheetUrl').value);
  if (!parsed) {
    setStatus('lookerStatus', 'Paste the full URL of the Google Sheet you imported the CSV into.', 'err');
    return;
  }
  const url = LookerStudio.buildReportUrl({
    templateReportId: template.lookerReportId,
    reportName: 'KPI Dashboard',
    spreadsheetId: parsed.spreadsheetId,
    gid: parsed.gid
  });
  window.open(url, '_blank');
  setStatus('lookerStatus', 'Opened Looker Studio in a new tab.', 'ok');
}

/* ------------------------------------------------------------------- boot */
Office.onReady(() => {
  wireStep1();
  wireStep2();
  wireStep3();
  wireStep4();
});
