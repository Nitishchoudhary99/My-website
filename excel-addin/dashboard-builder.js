/* ==========================================================================
   DocBrisk Dashboard Builder (Office.js)
   Takes the slots produced by KpiTemplates.autoFill() and writes a real,
   native Excel dashboard: a new sheet with scorecard tiles (styled cells,
   not images) and Excel-native charts built from a small hidden data
   block. Native charts because they're editable by the user afterwards
   with Excel's own chart tools — a picture of a chart would not be.
   ========================================================================== */

const DashboardBuilder = (() => {
  const SHEET_NAME = 'KPI Dashboard';
  const DATA_SHEET_NAME = 'KPI Dashboard Data';

  const NUMBER_FORMAT = {
    currency: '"₹"#,##0',
    percent: '0.0"%"',
    number: '#,##0'
  };

  async function build(context, filledSlots) {
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    await context.sync();

    // Remove a previous run's sheets so re-generating doesn't pile up
    // "KPI Dashboard (2)", "(3)" ... every time the user clicks Generate.
    for (const s of sheets.items) {
      if (s.name === SHEET_NAME || s.name === DATA_SHEET_NAME) s.delete();
    }
    await context.sync();

    const dataSheet = sheets.add(DATA_SHEET_NAME);
    const dashSheet = sheets.add(SHEET_NAME);
    dataSheet.visibility = Excel.SheetVisibility.hidden;

    let dataRow = 0;
    const scorecards = filledSlots.filter((s) => s.kpi && s.kpi.chart === 'scorecard');
    const charts = filledSlots.filter((s) => s.kpi && s.kpi.chart !== 'scorecard');

    layoutScorecards(dashSheet, scorecards);
    for (const slot of charts) {
      dataRow = await writeChartData(dataSheet, dataRow, slot.kpi);
    }
    await context.sync();

    let chartTop = 40 + Math.ceil(scorecards.length / 4) * 110 + 20;
    for (const slot of charts) {
      chartTop = await addChart(context, dashSheet, dataSheet, slot, chartTop);
    }

    dashSheet.getRange('A1').select();
    await context.sync();
  }

  function layoutScorecards(sheet, scorecardSlots) {
    const COLS_PER_ROW = 4;
    const TILE_W = 2; // columns spanned per tile
    sheet.getRange('A1:M1').format.rowHeight = 6; // top margin

    scorecardSlots.forEach((slot, i) => {
      const col = (i % COLS_PER_ROW) * (TILE_W + 1);
      const row = 1 + Math.floor(i / COLS_PER_ROW) * 5;
      const kpi = slot.kpi;

      const labelCell = sheet.getRangeByIndexes(row, col, 1, TILE_W + 1);
      const valueCell = sheet.getRangeByIndexes(row + 1, col, 2, TILE_W + 1);
      const noteCell = sheet.getRangeByIndexes(row + 3, col, 1, TILE_W + 1);

      labelCell.merge();
      valueCell.merge();
      noteCell.merge();

      labelCell.values = [[kpi.label.toUpperCase()]];
      labelCell.format.font.size = 9;
      labelCell.format.font.bold = true;
      labelCell.format.font.color = '#64748B';

      valueCell.values = [[formatValue(kpi)]];
      valueCell.format.font.size = 26;
      valueCell.format.font.bold = true;
      valueCell.format.font.color = '#312E81';
      valueCell.format.verticalAlignment = Excel.VerticalAlignment.center;

      noteCell.values = [[kpi.explain || '']];
      noteCell.format.font.size = 8;
      noteCell.format.font.italic = true;
      noteCell.format.font.color = '#94A3B8';

      const tile = sheet.getRangeByIndexes(row, col, 4, TILE_W + 1);
      tile.format.fill.color = '#F8FAFC';
      tile.format.borders.getItem('EdgeTop').style = Excel.BorderLineStyle.continuous;
      tile.format.borders.getItem('EdgeBottom').style = Excel.BorderLineStyle.continuous;
      tile.format.borders.getItem('EdgeLeft').style = Excel.BorderLineStyle.continuous;
      tile.format.borders.getItem('EdgeRight').style = Excel.BorderLineStyle.continuous;
      ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'].forEach((edge) => {
        tile.format.borders.getItem(edge).color = '#E2E8F0';
      });
    });
  }

  function formatValue(kpi) {
    if (kpi.format === 'currency') return '₹' + KpiEngine.formatNumber(kpi.value);
    if (kpi.format === 'percent') return KpiEngine.formatNumber(kpi.value) + '%';
    return KpiEngine.formatNumber(kpi.value);
  }

  /** Write one chart's series into the hidden data sheet; returns the next
   *  free row so charts don't overwrite each other's data. */
  async function writeChartData(dataSheet, startRow, kpi) {
    const series = kpi.buildSeries ? kpi.buildSeries() : [];
    kpi._dataRange = null;
    if (!series.length) return startRow;

    const headerRow = [[kpi.label, '']];
    const body = series.map((p) => [p.label, p.value]);
    const range = dataSheet.getRangeByIndexes(startRow, 0, body.length + 1, 2);
    range.values = [...headerRow, ...body];
    kpi._dataRangeAddress = { row: startRow, count: body.length + 1 };
    return startRow + body.length + 2; // blank row between blocks
  }

  async function addChart(context, dashSheet, dataSheet, slot, top) {
    const kpi = slot.kpi;
    if (!kpi._dataRangeAddress) return top;
    const { row, count } = kpi._dataRangeAddress;
    const range = dataSheet.getRangeByIndexes(row, 0, count, 2);

    const chartType = kpi.chart === 'line' ? Excel.ChartType.line : Excel.ChartType.columnClustered;
    const chart = dashSheet.charts.add(chartType, range, Excel.ChartSeriesBy.columns);
    chart.title.text = kpi.label;
    chart.legend.visible = false;
    chart.setPosition('A1', null);
    chart.top = top;
    chart.left = 20;
    chart.width = slot.size === 'wide' ? 640 : 300;
    chart.height = 260;
    await context.sync();
    return top + 280;
  }

  return { build, SHEET_NAME };
})();
