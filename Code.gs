/** Automated Sales Reporting — portfolio demonstration. */
const CONFIG = Object.freeze({
  raw: 'Raw_Data', clean: 'Clean_Data', summary: 'Summary', dashboard: 'Dashboard',
  headers: ['Order ID','Date','Product','Category','Quantity','Unit Price','Status']
});

const DASHBOARD_STYLE = Object.freeze({
  navy: '#17365D',
  green: '#1F8A70',
  greenLight: '#E8F5F1',
  ink: '#243447',
  muted: '#667085',
  canvas: '#F7F9FC',
  white: '#FFFFFF',
  border: '#D9E2EC',
  warning: '#B54708'
});

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Sales Automation')
    .addItem('Load sample data (demo only)', 'loadSampleData')
    .addItem('Refresh reports', 'refreshReports')
    .addItem('Enable daily refresh', 'enableDailyRefresh')
    .addItem('Disable daily refresh', 'disableDailyRefresh')
    .addToUi();
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function loadSampleData() {
  const sh = sheet_(CONFIG.raw);
  if (sh.getLastRow() > 1) {
    const choice = SpreadsheetApp.getUi().alert('Replace data?',
      'This will overwrite Raw_Data. Continue only with demo data.', SpreadsheetApp.getUi().ButtonSet.YES_NO);
    if (choice !== SpreadsheetApp.getUi().Button.YES) return;
  }
  const products = [
    ['Laptop Stand','Accessories',35],['Wireless Mouse','Accessories',22],
    ['USB-C Hub','Electronics',49],['Webcam','Electronics',75],
    ['Desk Lamp','Office',40],['Notebook','Office',12]
  ];
  const rows = [CONFIG.headers];
  for (let i = 0; i < 120; i++) {
    const p = products[i % products.length];
    const d = new Date(2026, Math.floor(i / 20), (i % 20) + 1, 12);
    rows.push(['ORD-' + String(i + 1).padStart(4, '0'), d, p[0], p[1], (i % 4) + 1, p[2], i % 11 === 0 ? 'Cancelled' : 'Completed']);
  }
  sh.clear();
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.getRange(2, 2, rows.length - 1, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, 6, rows.length - 1, 1).setNumberFormat('$#,##0.00');
  formatHeader_(sh, 7);
  refreshReports();
}

function refreshReports() {
  const raw = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.raw);
  if (!raw || raw.getLastRow() < 1) throw new Error('Raw_Data must contain the required headers.');
  const values = raw.getDataRange().getValues();
  const header = values[0].map(v => String(v).trim().toLowerCase());
  const indices = CONFIG.headers.map(h => header.indexOf(h.toLowerCase()));
  if (indices.some(i => i < 0)) throw new Error('Raw_Data is missing required columns: ' + CONFIG.headers.join(', '));
  const seen = new Set();
  const clean = [];
  let invalid = 0, duplicates = 0, excluded = 0;
  for (const original of values.slice(1)) {
    const r = indices.map(i => original[i]);
    const id = String(r[0]).trim();
    const date = parseDate_(r[1]);
    const product = String(r[2]).trim();
    const category = String(r[3]).trim();
    const qty = Number(r[4]);
    const price = Number(r[5]);
    const status = String(r[6]).trim().toLowerCase();
    if (!id || !date || !product || !category || !Number.isFinite(qty) || !Number.isFinite(price) || qty <= 0 || !Number.isInteger(qty) || price < 0) { invalid++; continue; }
    if (seen.has(id)) { duplicates++; continue; }
    seen.add(id);
    if (status !== 'completed') { excluded++; continue; }
    clean.push([id, date, product, category, qty, price, qty * price]);
  }
  clean.sort((a,b) => a[1] - b[1]);
  const total = clean.reduce((sum,r) => sum + r[6], 0);
  const monthly = new Map(), categories = new Map();
  for (const r of clean) {
    const month = Utilities.formatDate(r[1], Session.getScriptTimeZone(), 'yyyy-MM');
    monthly.set(month, (monthly.get(month) || 0) + r[6]);
    categories.set(r[3], (categories.get(r[3]) || 0) + r[6]);
  }
  const monthRows = [...monthly.entries()].sort((a,b) => a[0].localeCompare(b[0]));
  const categoryRows = [...categories.entries()].sort((a,b) => b[1] - a[1]);
  validateReportCalculations_(clean, total, invalid, duplicates, excluded,
    values.length - 1, monthRows, categoryRows);
  const cleanSheet = sheet_(CONFIG.clean);
  cleanSheet.clear();
  const cleanedRows = [['Order ID','Date','Product','Category','Quantity','Unit Price','Revenue'], ...clean];
  cleanSheet.getRange(1, 1, cleanedRows.length, 7).setValues(cleanedRows);
  formatHeader_(cleanSheet, 7);
  if (clean.length) {
    cleanSheet.getRange(2,2,clean.length,1).setNumberFormat('yyyy-mm-dd');
    cleanSheet.getRange(2,6,clean.length,2).setNumberFormat('$#,##0.00');
  }
  const summary = sheet_(CONFIG.summary);
  summary.clear();
  const summaryRows = [
    ['Metric','Value','','Monthly Sales','Revenue','','Category','Revenue'],
    ['Total Sales',total,'',...(monthRows[0] || ['', '']),'',...(categoryRows[0] || ['', ''])],
    ['Orders',clean.length,'',...(monthRows[1] || ['', '']),'',...(categoryRows[1] || ['', ''])],
    ['Average Order',clean.length ? total/clean.length : 0,'',...(monthRows[2] || ['', '']),'',...(categoryRows[2] || ['', ''])],
    ['Invalid Rows',invalid,'',...(monthRows[3] || ['', '']),'',...(categoryRows[3] || ['', ''])],
    ['Duplicate IDs',duplicates,'',...(monthRows[4] || ['', '']),'',...(categoryRows[4] || ['', ''])],
    ['Excluded Status',excluded,'',...(monthRows[5] || ['', '']),'',...(categoryRows[5] || ['', ''])]
  ];
  const max = Math.max(summaryRows.length, monthRows.length+1, categoryRows.length+1);
  while (summaryRows.length < max) summaryRows.push(['','','','','','','','']);
  monthRows.forEach((r,i) => { summaryRows[i+1][3]=r[0]; summaryRows[i+1][4]=r[1]; });
  categoryRows.forEach((r,i) => { summaryRows[i+1][6]=r[0]; summaryRows[i+1][7]=r[1]; });
  summary.getRange(1,1,summaryRows.length,8).setValues(summaryRows);
  formatHeader_(summary,8);
  summary.getRange('B2').setNumberFormat('$#,##0.00');
  summary.getRange('B4').setNumberFormat('$#,##0.00');
  if (monthRows.length) summary.getRange(2,5,monthRows.length,1).setNumberFormat('$#,##0.00');
  if (categoryRows.length) summary.getRange(2,8,categoryRows.length,1).setNumberFormat('$#,##0.00');
  buildDashboard_(total, clean.length, invalid, duplicates, excluded, monthRows, categoryRows);
  SpreadsheetApp.flush();
}

function parseDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v.trim())) {
    const [y,m,d] = v.trim().split('-').map(Number);
    const date = new Date(y,m-1,d,12);
    if (date.getFullYear()===y && date.getMonth()===m-1 && date.getDate()===d) return date;
  }
  return null;
}

function formatHeader_(sh, count) {
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,count).setBackground('#17365d').setFontColor('#ffffff').setFontWeight('bold');
  sh.autoResizeColumns(1,count);
}

/**
 * Guards the established report calculations against accidental regressions.
 * These checks do not alter data; refresh stops before replacing report output
 * if row accounting or an aggregate no longer reconciles.
 */
function validateReportCalculations_(clean, total, invalid, duplicates, excluded,
    sourceRowCount, monthly, categories) {
  const tolerance = Math.max(0.01, Math.abs(total) * 1e-9);
  const calculatedTotal = clean.reduce((sum, row) => sum + row[6], 0);
  const monthlyTotal = monthly.reduce((sum, row) => sum + row[1], 0);
  const categoryTotal = categories.reduce((sum, row) => sum + row[1], 0);
  const accountedRows = clean.length + invalid + duplicates + excluded;
  const allCountsValid = [invalid, duplicates, excluded, clean.length]
    .every(value => Number.isInteger(value) && value >= 0);

  if (!allCountsValid) throw new Error('Regression check failed: invalid report counts.');
  if (accountedRows !== sourceRowCount) {
    throw new Error('Regression check failed: source-row accounting does not reconcile.');
  }
  if (Math.abs(calculatedTotal - total) > tolerance ||
      Math.abs(monthlyTotal - total) > tolerance ||
      Math.abs(categoryTotal - total) > tolerance) {
    throw new Error('Regression check failed: revenue aggregates do not reconcile.');
  }
}

function buildDashboard_(total, orders, invalid, duplicates, excluded, monthly, categories) {
  const sh = sheet_(CONFIG.dashboard);
  resetDashboard_(sh);
  formatDashboardCanvas_(sh);
  renderDashboardHeader_(sh);
  renderKpiCards_(sh, total, orders);
  renderDataQuality_(sh, invalid, duplicates, excluded);
  renderDashboardCharts_(sh, monthly.length, categories.length);
}

function resetDashboard_(sh) {
  // Chart removal first guarantees that repeated refreshes never accumulate charts.
  sh.getCharts().forEach(chart => sh.removeChart(chart));
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.setConditionalFormatRules([]);
}

function formatDashboardCanvas_(sh) {
  const style = DASHBOARD_STYLE;
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(3);
  sh.getRange('A1:L29').setBackground(style.canvas).setFontFamily('Arial')
    .setFontColor(style.ink).setVerticalAlignment('middle');
  sh.setColumnWidths(1, 12, 80); // 960 px keeps the dashboard compact at typical laptop widths.
  sh.setRowHeights(1, 29, 22);
  sh.setRowHeight(1, 34);
  sh.setRowHeight(2, 28);
  sh.setRowHeight(3, 12);
  sh.setRowHeights(5, 4, 28);
  sh.setRowHeight(9, 12);
  sh.setRowHeights(10, 3, 24);
  sh.setRowHeight(13, 12);
  if (sh.getMaxColumns() > 12) sh.hideColumns(13, sh.getMaxColumns() - 12);
  if (sh.getMaxRows() > 29) sh.hideRows(30, sh.getMaxRows() - 29);
}

function renderDashboardHeader_(sh) {
  const style = DASHBOARD_STYLE;
  sh.getRange('A1:L1').merge().setValue('SALES PERFORMANCE DASHBOARD')
    .setBackground(style.navy).setFontColor(style.white).setFontSize(19)
    .setFontWeight('bold').setHorizontalAlignment('left');
  sh.getRange('A2:L2').merge()
    .setValue('Revenue, order performance, and data quality  |  Fictional portfolio data')
    .setBackground(style.white).setFontColor(style.muted).setFontSize(10)
    .setHorizontalAlignment('left');
}

function renderKpiCards_(sh, total, orders) {
  const average = orders ? total / orders : 0;
  renderKpiCard_(sh, 'A5:D8', 'TOTAL REVENUE', total, '$#,##0.00');
  renderKpiCard_(sh, 'E5:H8', 'COMPLETED ORDERS', orders, '#,##0');
  renderKpiCard_(sh, 'I5:L8', 'AVERAGE ORDER VALUE', average, '$#,##0.00');
}

function renderKpiCard_(sh, a1Notation, label, value, numberFormat) {
  const style = DASHBOARD_STYLE;
  const card = sh.getRange(a1Notation);
  const row = card.getRow(), column = card.getColumn(), columns = card.getNumColumns();
  card.setBackground(style.white).setBorder(true, true, true, true, false, false,
    style.border, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row, column, 1, columns).merge().setValue(label)
    .setFontColor(style.navy).setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('center');
  sh.getRange(row + 1, column, 3, columns).merge().setValue(value)
    .setNumberFormat(numberFormat).setFontColor(style.green).setFontSize(22)
    .setFontWeight('bold').setHorizontalAlignment('center');
}

function renderDataQuality_(sh, invalid, duplicates, excluded) {
  const style = DASHBOARD_STYLE;
  sh.getRange('A10:C12').merge().setValue('DATA QUALITY')
    .setBackground(style.navy).setFontColor(style.white).setFontWeight('bold')
    .setHorizontalAlignment('center');
  const metrics = [
    ['D10:F10', 'D11:F12', 'INVALID ROWS', invalid],
    ['G10:I10', 'G11:I12', 'DUPLICATE IDS', duplicates],
    ['J10:L10', 'J11:L12', 'EXCLUDED RECORDS', excluded]
  ];
  metrics.forEach(metric => {
    sh.getRange(metric[0]).merge().setValue(metric[2]).setBackground(style.white)
      .setFontColor(style.muted).setFontSize(9).setFontWeight('bold')
      .setHorizontalAlignment('center');
    sh.getRange(metric[1]).merge().setValue(metric[3]).setBackground(style.white)
      .setFontColor(metric[3] === 0 ? style.green : style.warning).setFontSize(16)
      .setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange(metric[0].split(':')[0] + ':' + metric[1].split(':')[1])
      .setBorder(true, true, true, true, false, false, style.border,
        SpreadsheetApp.BorderStyle.SOLID);
  });
}

function renderDashboardCharts_(sh, monthCount, categoryCount) {
  const summary = sheet_(CONFIG.summary);
  if (monthCount) {
    sh.insertChart(sh.newChart().setChartType(Charts.ChartType.COLUMN)
      .addRange(summary.getRange(1, 4, monthCount + 1, 2)).setPosition(14, 1, 0, 0)
      .setOption('title', 'Monthly Revenue').setOption('width', 470).setOption('height', 315)
      .setOption('backgroundColor', DASHBOARD_STYLE.white)
      .setOption('colors', [DASHBOARD_STYLE.green]).setOption('legend', {position: 'none'})
      .setOption('titleTextStyle', {color: DASHBOARD_STYLE.navy, fontSize: 14, bold: true})
      .setOption('hAxis', {textStyle: {color: DASHBOARD_STYLE.muted}})
      .setOption('vAxis', {format: '$#,##0', minValue: 0, textStyle: {color: DASHBOARD_STYLE.muted}})
      .build());
  } else {
    renderEmptyChartState_(sh, 'A14:F28', 'Monthly Revenue', 'No completed orders to chart');
  }
  if (categoryCount) {
    sh.insertChart(sh.newChart().setChartType(Charts.ChartType.PIE)
      .addRange(summary.getRange(1, 7, categoryCount + 1, 2)).setPosition(14, 7, 0, 0)
      .setOption('title', 'Revenue by Category').setOption('width', 470).setOption('height', 315)
      .setOption('pieHole', 0.55).setOption('pieSliceText', 'percentage')
      .setOption('backgroundColor', DASHBOARD_STYLE.white)
      .setOption('colors', ['#17365D', '#1F8A70', '#5B8FF9', '#A5D6C8', '#8C9EFF', '#7A869A'])
      .setOption('titleTextStyle', {color: DASHBOARD_STYLE.navy, fontSize: 14, bold: true})
      .setOption('legend', {position: 'right', textStyle: {color: DASHBOARD_STYLE.muted, fontSize: 10}})
      .build());
  } else {
    renderEmptyChartState_(sh, 'G14:L28', 'Revenue by Category', 'No completed orders to chart');
  }
}

function renderEmptyChartState_(sh, a1Notation, title, message) {
  const range = sh.getRange(a1Notation);
  range.merge().setValue(title + '\n\n' + message).setWrap(true)
    .setBackground(DASHBOARD_STYLE.white).setFontColor(DASHBOARD_STYLE.muted)
    .setFontSize(12).setHorizontalAlignment('center')
    .setBorder(true, true, true, true, false, false, DASHBOARD_STYLE.border,
      SpreadsheetApp.BorderStyle.SOLID);
}

function enableDailyRefresh() {
  disableDailyRefresh();
  ScriptApp.newTrigger('refreshReports').timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getUi().alert('Daily refresh enabled (approximately 7 AM in script timezone).');
}
function disableDailyRefresh() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction()==='refreshReports')
    .forEach(t => ScriptApp.deleteTrigger(t));
}
