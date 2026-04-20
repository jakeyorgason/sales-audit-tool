function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const sharedSecret = PropertiesService.getScriptProperties().getProperty('SHARED_SECRET') || '';
    if (sharedSecret && data.sharedSecret !== sharedSecret) {
      throw new Error('Unauthorized request. Shared secret mismatch.');
    }

    const templateId = data.templateId;
    const destinationFolderId = data.destinationFolderId;
    const reportName = data.reportName || 'Sales Audit Report';
    const brandName = data.brandName || '';
    const dateRangeLabel = data.dateRangeLabel || '';

    const kpiSummary = data.kpiSummary || {};
    const wasteSummary = data.wasteSummary || {};
    const matchTypeRevenueRows = data.matchTypeRevenueRows || [];
    const matchTypeInefficientRows = data.matchTypeInefficientRows || [];
    const campaignRows = data.campaignRows || [];
    const campaignTypeRows = data.campaignTypeRows || [];
    const topKeywordRows = data.topKeywordRows || [];
    const topSearchTermRows = data.topSearchTermRows || [];
    const wasteKeywordRows = data.wasteKeywordRows || [];
    const wasteSearchTermRows = data.wasteSearchTermRows || [];
    const winnerKeywordRows = data.winnerKeywordRows || [];
    const winnerSearchTermRows = data.winnerSearchTermRows || [];
    const targetingDataRows = data.targetingDataRows || [];
    const searchTermDataRows = data.searchTermDataRows || [];
    const lead = data.lead || {};

    if (!templateId || !destinationFolderId) {
      throw new Error('Missing template or destination folder ID.');
    }

    const templateFile = DriveApp.getFileById(templateId);
    const destinationFolder = DriveApp.getFolderById(destinationFolderId);
    const copiedFile = templateFile.makeCopy(reportName, destinationFolder);
    const spreadsheet = SpreadsheetApp.openById(copiedFile.getId());
    const summarySheet = spreadsheet.getSheetByName('Summary');
    if (!summarySheet) {
      throw new Error('Summary tab not found.');
    }

    writeSummaryTab_(spreadsheet, brandName, dateRangeLabel, kpiSummary, wasteSummary);
    writeMatchTypeSections_(spreadsheet, matchTypeRevenueRows, matchTypeInefficientRows);
    writeCampaignChartData_(spreadsheet, campaignRows, campaignTypeRows);
    writeTopSpendersTab_(spreadsheet, topKeywordRows, topSearchTermRows);
    writeWasteTab_(spreadsheet, wasteKeywordRows, wasteSearchTermRows);
    writeWinnersTab_(spreadsheet, winnerKeywordRows, winnerSearchTermRows);
    writeTargetingDataTab_(spreadsheet, targetingDataRows);
    writeSearchTermDataTab_(spreadsheet, searchTermDataRows);

    buildOverallMatchTypePerformanceChart_(summarySheet);
    buildMatchTypeInefficientSpendChart_(summarySheet);
    buildMatchTypeSalesDistributionChart_(summarySheet);
    buildMatchTypeInefficientSpendDistributionChart_(summarySheet);
    buildCampaignSalesDistributionChart_(summarySheet);
    buildCampaignSpendDistributionChart_(summarySheet);
    buildCampaignTypeSalesDistributionChart_(summarySheet);

    SpreadsheetApp.flush();

    maybeLogLead_(lead, spreadsheet.getUrl(), copiedFile.getName());

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      fileId: copiedFile.getId(),
      url: spreadsheet.getUrl(),
      name: copiedFile.getName()
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: String(error)
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function maybeLogLead_(lead, reportUrl, reportName) {
  const logSheetId = PropertiesService.getScriptProperties().getProperty('LEAD_LOG_SHEET_ID') || '';
  if (!logSheetId || !lead) return;

  const ss = SpreadsheetApp.openById(logSheetId);
  const sheet = ss.getSheetByName('Leads') || ss.insertSheet('Leads');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Timestamp', 'Name', 'Email', 'Phone', 'Brand Name', 'Source', 'Report Name', 'Report URL']);
  }
  sheet.appendRow([
    new Date(),
    lead.name || '',
    lead.email || '',
    lead.phone || '',
    lead.brandName || '',
    lead.source || '',
    reportName || '',
    reportUrl || ''
  ]);
}

function writeSummaryTab_(spreadsheet, brandName, dateRangeLabel, kpiSummary, wasteSummary) {
  const sheet = spreadsheet.getSheetByName('Summary');
  sheet.getRange('E2').setValue(brandName);
  sheet.getRange('B3').setValue(dateRangeLabel);

  const totalSales = Number(kpiSummary.total_sales || 0);
  const adSales = Number(kpiSummary.ad_sales || 0);
  const organicSales = Number(kpiSummary.organic_sales || 0);
  const ppcSalesShare = totalSales > 0 ? adSales / totalSales : 0;
  const organicSalesShare = totalSales > 0 ? organicSales / totalSales : 0;

  sheet.getRange('D6').setValue(ppcSalesShare);
  sheet.getRange('E6').setValue(organicSalesShare);
  sheet.getRange('D6:E6').setNumberFormat('0.00%');

  const values = [
    [toPercent_(kpiSummary.tacos_pct)],
    [toCurrency_(kpiSummary.ad_sales)],
    [toCurrency_(kpiSummary.organic_sales)],
    [toCurrency_(wasteSummary.wasted_spend)],
    [toCurrency_(kpiSummary.spend)],
    [toCurrency_(kpiSummary.total_sales)],
    [toPercent_(kpiSummary.acos_pct)],
    [toNumber_(kpiSummary.sessions)],
    [toPercent_(kpiSummary.unit_session_percentage)],
    [toNumber_(kpiSummary.units_ordered)],
    [toNumber_(kpiSummary.roas)],
    [toCurrency_(kpiSummary.ntb_sales)],
    [toNumber_(kpiSummary.ntb_orders)],
    [toPercent_(kpiSummary.ntb_sales_pct)],
    [toPercent_(kpiSummary.ntb_orders_pct)],
    [toCurrency_(wasteSummary.spend_no_sale)]
  ];

  const valueRange = sheet.getRange('B5:B20');
  valueRange.setValues(values);
  sheet.getRange('B5').setNumberFormat('0.00%');
  sheet.getRange('B6:B10').setNumberFormat('$#,##0.00');
  sheet.getRange('B11').setNumberFormat('0.00%');
  sheet.getRange('B12').setNumberFormat('#,##0');
  sheet.getRange('B13').setNumberFormat('0.00%');
  sheet.getRange('B14').setNumberFormat('#,##0');
  sheet.getRange('B15').setNumberFormat('0.00');
  sheet.getRange('B16').setNumberFormat('$#,##0.00');
  sheet.getRange('B17').setNumberFormat('#,##0');
  sheet.getRange('B18:B19').setNumberFormat('0.00%');
  sheet.getRange('B20').setNumberFormat('$#,##0.00');
}

function writeMatchTypeSections_(spreadsheet, revenueRows, inefficientRows) {
  const sheet = spreadsheet.getSheetByName('Summary');
  const revenueMatrix = buildMatchTypeMatrix_(revenueRows, true);
  const inefficientMatrix = buildMatchTypeMatrix_(inefficientRows, true);
  sheet.getRange('A24:J30').setValues(revenueMatrix);
  sheet.getRange('A33:J39').setValues(inefficientMatrix);
  sheet.getRange('A24:J24').setFontWeight('bold');
  sheet.getRange('A33:J33').setFontWeight('bold');
  sheet.getRange('B25:B30').setNumberFormat('#,##0');
  sheet.getRange('C25:C30').setNumberFormat('#,##0');
  sheet.getRange('D25:E30').setNumberFormat('$#,##0.00');
  sheet.getRange('B34:B39').setNumberFormat('#,##0');
  sheet.getRange('C34:C39').setNumberFormat('#,##0');
  sheet.getRange('D34:E39').setNumberFormat('$#,##0.00');
  sheet.getRange('F25:J30').setNumberFormat('0.00%');
  sheet.getRange('F34:J39').setNumberFormat('0.00%');
  sheet.getRange('A24:J39').setHorizontalAlignment('center');
}

function buildMatchTypeMatrix_(rows, includeBrandedKw) {
  const rowOrder = includeBrandedKw ? ['AUTO', 'BROAD', 'EXACT', 'PHRASE', 'Grand Total', 'Branded KW'] : ['AUTO', 'BROAD', 'EXACT', 'PHRASE', 'Grand Total'];
  const rowMap = {};
  (rows || []).forEach(r => { rowMap[normalizeMatchTypeLabel_(r.match_type || r.matchType || r.label || '')] = r; });
  const totals = { impressions: 0, clicks: 0, spend: 0, sales: 0 };
  ['AUTO', 'BROAD', 'EXACT', 'PHRASE'].forEach(mt => {
    const row = rowMap[mt] || {};
    totals.impressions += Number(row.impressions || 0);
    totals.clicks += Number(row.clicks || 0);
    totals.spend += Number(row.spend || 0);
    totals.sales += Number(row.sales || 0);
  });
  const output = [['Match Type', 'Impressions', 'Clicks', 'Spend', 'Sales', 'Impressions', 'Clicks', 'Spend', 'Sales', 'ACoS']];
  rowOrder.forEach(label => {
    const row = label === 'Grand Total' ? totals : (rowMap[label] || {});
    const impressions = Number(row.impressions || 0);
    const clicks = Number(row.clicks || 0);
    const spend = Number(row.spend || 0);
    const sales = Number(row.sales || 0);
    const pctImpressions = totals.impressions > 0 ? impressions / totals.impressions : 0;
    const pctClicks = totals.clicks > 0 ? clicks / totals.clicks : 0;
    const pctSpend = totals.spend > 0 ? spend / totals.spend : 0;
    const pctSales = totals.sales > 0 ? sales / totals.sales : 0;
    const acos = sales > 0 ? spend / sales : 0;
    output.push([label, impressions, clicks, spend, sales, pctImpressions, pctClicks, pctSpend, pctSales, acos]);
  });
  return output;
}

function writeCampaignChartData_(spreadsheet, campaignRows, campaignTypeRows) {
  const sheet = spreadsheet.getSheetByName('Summary');
  const topCampaignsBySales = (campaignRows || []).slice().sort((a, b) => Number(b.sales || 0) - Number(a.sales || 0)).slice(0, 10);
  const topCampaignsBySpend = (campaignRows || []).slice().sort((a, b) => Number(b.spend || 0) - Number(a.spend || 0)).slice(0, 10);
  const salesChartRows = topCampaignsBySales.map(r => [String(r.campaign_name || ''), Number(r.sales || 0)]);
  const spendChartRows = topCampaignsBySpend.map(r => [String(r.campaign_name || ''), Number(r.spend || 0)]);
  const typeChartRows = (campaignTypeRows || []).map(r => [String(r.campaign_type || ''), Number(r.sales || 0)]);
  sheet.getRange('A80:B117').clearContent();
  sheet.getRange('E80:F98').clearContent();
  if (salesChartRows.length) sheet.getRange(80, 1, salesChartRows.length, 2).setValues(salesChartRows);
  if (spendChartRows.length) sheet.getRange(80, 5, spendChartRows.length, 2).setValues(spendChartRows);
  if (typeChartRows.length) sheet.getRange(99, 1, typeChartRows.length, 2).setValues(typeChartRows);
}

function buildOverallMatchTypePerformanceChart_(sheet) {
  removeChartAt_(sheet, 43, 1);
  const chart = sheet.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange('A24:A28')).addRange(sheet.getRange('D24:E28'))
    .setOption('title', 'Overall Match Type Performance')
    .setOption('legend', { position: 'top' })
    .setOption('hAxis', { title: 'Match Type' })
    .setOption('vAxis', { title: 'Dollar amount', format: '$#,##0' })
    .setOption('colors', ['#415A68', '#F47322']).setPosition(43, 1, 0, 0).setNumHeaders(1).build();
  sheet.insertChart(chart);
}

function buildMatchTypeInefficientSpendChart_(sheet) {
  removeChartAt_(sheet, 43, 5);
  const chart = sheet.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange('A33:A37')).addRange(sheet.getRange('D33:E37'))
    .setOption('title', 'Match Type Inefficient Spend Data')
    .setOption('legend', { position: 'top' })
    .setOption('hAxis', { title: 'Match Type' })
    .setOption('vAxis', { title: 'Dollar amount', format: '$#,##0' })
    .setOption('colors', ['#415A68', '#F47322']).setPosition(43, 5, 0, 0).setNumHeaders(1).build();
  sheet.insertChart(chart);
}

function buildMatchTypeSalesDistributionChart_(sheet) {
  removeChartAt_(sheet, 62, 1);
  const chart = sheet.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange('A25:A28')).addRange(sheet.getRange('I25:I28'))
    .setOption('title', 'Match Type Sales Distribution')
    .setOption('legend', { position: 'left' }).setOption('pieHole', 0.45)
    .setOption('pieSliceText', 'percentage').setOption('colors', ['#415A68', '#F47322', '#FDBA31'])
    .setPosition(62, 1, 0, 0).setNumHeaders(0).build();
  sheet.insertChart(chart);
}

function buildMatchTypeInefficientSpendDistributionChart_(sheet) {
  removeChartAt_(sheet, 62, 5);
  const chart = sheet.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange('A34:A37')).addRange(sheet.getRange('H34:H37'))
    .setOption('title', 'Match Type - Inefficient Ad Spend')
    .setOption('legend', { position: 'left' }).setOption('pieHole', 0.45)
    .setOption('pieSliceText', 'percentage').setOption('colors', ['#415A68', '#F47322', '#FDBA31'])
    .setPosition(62, 5, 0, 0).setNumHeaders(0).build();
  sheet.insertChart(chart);
}

function buildCampaignSalesDistributionChart_(sheet) {
  removeChartAt_(sheet, 80, 1);
  const chart = sheet.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange('A80:B89')).setOption('title', 'Campaign - Sales Distribution')
    .setOption('legend', { position: 'left' }).setOption('pieSliceText', 'percentage')
    .setOption('colors', ['#415A68', '#F47322', '#FDBA31']).setPosition(80, 1, 0, 0).setNumHeaders(0).build();
  sheet.insertChart(chart);
}

function buildCampaignSpendDistributionChart_(sheet) {
  removeChartAt_(sheet, 80, 5);
  const chart = sheet.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange('E80:F89')).setOption('title', 'Campaign - Spend Distribution')
    .setOption('legend', { position: 'left' }).setOption('pieSliceText', 'percentage')
    .setOption('colors', ['#415A68', '#F47322', '#FDBA31']).setPosition(80, 5, 0, 0).setNumHeaders(0).build();
  sheet.insertChart(chart);
}

function buildCampaignTypeSalesDistributionChart_(sheet) {
  removeChartAt_(sheet, 99, 1);
  const chart = sheet.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange('A99:B102')).setOption('title', 'Campaign Type - Sales Distribution')
    .setOption('legend', { position: 'left' }).setOption('pieSliceText', 'percentage')
    .setOption('colors', ['#415A68', '#F47322']).setPosition(99, 1, 0, 0).setNumHeaders(0).build();
  sheet.insertChart(chart);
}

function removeChartAt_(sheet, anchorRow, anchorColumn) {
  sheet.getCharts().forEach(chart => {
    const info = chart.getContainerInfo();
    if (info && info.getAnchorRow() === anchorRow && info.getAnchorColumn() === anchorColumn) {
      sheet.removeChart(chart);
    }
  });
}

function normalizeMatchTypeLabel_(value) {
  const v = String(value || '').trim().toUpperCase();
  if (v === 'AUTO' || v === 'AUTOMATIC') return 'AUTO';
  if (v === 'BROAD') return 'BROAD';
  if (v === 'EXACT') return 'EXACT';
  if (v === 'PHRASE') return 'PHRASE';
  if (v === 'GRAND TOTAL' || v === 'TOTAL') return 'Grand Total';
  if (v === 'BRANDED KW' || v === 'BRANDED' || v === 'BRANDED KEYWORD') return 'Branded KW';
  return v;
}

function toCurrency_(value) { return value === null || value === undefined || value === '' ? '' : Number(value); }
function toPercent_(value) { return value === null || value === undefined || value === '' ? '' : Number(value) / 100; }
function toNumber_(value) { return value === null || value === undefined || value === '' ? '' : Number(value); }
function getOrCreateSheet_(spreadsheet, name) { return spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name); }
function clearAndPrepSheet_(sheet) { sheet.clearContents(); sheet.clearFormats(); sheet.setFrozenRows(0); sheet.setFrozenColumns(0); }

function writeTableWithTitle_(sheet, startRow, startCol, title, rows) {
  sheet.getRange(startRow, startCol).setValue(title).setFontWeight('bold').setFontSize(12);
  if (!rows || !rows.length) {
    sheet.getRange(startRow + 1, startCol).setValue('No data');
    return;
  }
  const headers = Object.keys(rows[0]);
  sheet.getRange(startRow + 1, startCol, 1, headers.length).setValues([headers]).setFontWeight('bold');
  const values = rows.map(row => headers.map(h => row[h]));
  sheet.getRange(startRow + 2, startCol, values.length, headers.length).setValues(values);
}

function autoFormatSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;
  sheet.autoResizeColumns(1, lastCol);
  sheet.getRange(1, 1, lastRow, lastCol).setVerticalAlignment('middle');
}

function writeTopSpendersTab_(spreadsheet, topKeywordRows, topSearchTermRows) {
  const sheet = getOrCreateSheet_(spreadsheet, 'Top Spenders');
  clearAndPrepSheet_(sheet);
  writeTableWithTitle_(sheet, 1, 1, 'Top Keyword / Target Spenders', topKeywordRows);
  writeTableWithTitle_(sheet, 1, 8, 'Top Customer Search Term Spenders', topSearchTermRows);
  formatTwoTableSheet_(sheet); autoFormatSheet_(sheet);
}

function writeWasteTab_(spreadsheet, wasteKeywordRows, wasteSearchTermRows) {
  const sheet = getOrCreateSheet_(spreadsheet, 'Waste');
  clearAndPrepSheet_(sheet);
  writeTableWithTitle_(sheet, 1, 1, 'Keyword / Target Waste', wasteKeywordRows);
  writeTableWithTitle_(sheet, 1, 8, 'Customer Search Term Waste', wasteSearchTermRows);
  appendTotalsRowForSimpleTable_(sheet, 1, 1, wasteKeywordRows);
  appendTotalsRowForSimpleTable_(sheet, 1, 8, wasteSearchTermRows);
  formatTwoTableSheet_(sheet); autoFormatSheet_(sheet);
}

function writeWinnersTab_(spreadsheet, winnerKeywordRows, winnerSearchTermRows) {
  const sheet = getOrCreateSheet_(spreadsheet, 'Winners');
  clearAndPrepSheet_(sheet);
  writeTableWithTitle_(sheet, 1, 1, 'Winning Keyword / Targets', winnerKeywordRows);
  writeTableWithTitle_(sheet, 1, 8, 'Winning Customer Search Terms', winnerSearchTermRows);
  appendTotalsRowForSimpleTable_(sheet, 1, 1, winnerKeywordRows);
  appendTotalsRowForSimpleTable_(sheet, 1, 8, winnerSearchTermRows);
  formatTwoTableSheet_(sheet); autoFormatSheet_(sheet);
}

function writeTargetingDataTab_(spreadsheet, targetingDataRows) {
  const sheet = getOrCreateSheet_(spreadsheet, 'Targeting Data');
  clearAndPrepSheet_(sheet);
  writeTableWithTitle_(sheet, 1, 1, 'Targeting Data', targetingDataRows);
  formatPercentColumnsByHeader_(sheet); autoFormatSheet_(sheet);
}

function writeSearchTermDataTab_(spreadsheet, searchTermDataRows) {
  const sheet = getOrCreateSheet_(spreadsheet, 'Search Term Data');
  clearAndPrepSheet_(sheet);
  writeTableWithTitle_(sheet, 1, 1, 'Search Term Data', searchTermDataRows);
  formatPercentColumnsByHeader_(sheet); autoFormatSheet_(sheet);
}

function formatPercentColumnsByHeader_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;
  const headerRowValues = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  headerRowValues.forEach((header, idx) => {
    const h = String(header || '').trim().toLowerCase();
    const col = idx + 1;
    if (['ctr', 'cvr', 'acos', 'acos_pct'].includes(h)) sheet.getRange(3, col, Math.max(lastRow - 2, 1), 1).setNumberFormat('0.00%');
    if (['cpc'].includes(h)) sheet.getRange(3, col, Math.max(lastRow - 2, 1), 1).setNumberFormat('$#,##0.00');
    if (['roas'].includes(h)) sheet.getRange(3, col, Math.max(lastRow - 2, 1), 1).setNumberFormat('0.00');
    if (['spend', 'sales', 'ad_sales', 'total_sales', 'organic_sales'].includes(h)) sheet.getRange(3, col, Math.max(lastRow - 2, 1), 1).setNumberFormat('$#,##0.00');
    if (['impressions', 'clicks', 'orders', 'units_ordered', 'sessions'].includes(h)) sheet.getRange(3, col, Math.max(lastRow - 2, 1), 1).setNumberFormat('#,##0');
  });
}

function formatTwoTableSheet_(sheet) {
  const lastRow = sheet.getLastRow(); const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;
  const leftHeaders = sheet.getRange(2, 1, 1, Math.min(6, lastCol)).getValues()[0];
  leftHeaders.forEach((header, idx) => {
    const h = String(header || '').trim().toLowerCase(); const col = idx + 1;
    if (['acos'].includes(h)) sheet.getRange(3, col, Math.max(lastRow - 2, 1), 1).setNumberFormat('0.00%');
    if (['spend', 'sales'].includes(h)) sheet.getRange(3, col, Math.max(lastRow - 2, 1), 1).setNumberFormat('$#,##0.00');
  });
  if (lastCol >= 8) {
    const rightHeaders = sheet.getRange(2, 8, 1, lastCol - 7).getValues()[0];
    rightHeaders.forEach((header, idx) => {
      const h = String(header || '').trim().toLowerCase(); const col = idx + 8;
      if (['acos'].includes(h)) sheet.getRange(3, col, Math.max(lastRow - 2, 1), 1).setNumberFormat('0.00%');
      if (['spend', 'sales'].includes(h)) sheet.getRange(3, col, Math.max(lastRow - 2, 1), 1).setNumberFormat('$#,##0.00');
    });
  }
}

function appendTotalsRowForSimpleTable_(sheet, startRow, startCol, rows) {
  if (!rows || !rows.length) return;
  const headers = Object.keys(rows[0]);
  const totalsRowIndex = startRow + 2 + rows.length;
  const totalRow = new Array(headers.length).fill('');
  const termIdx = headers.findIndex(h => String(h).toLowerCase().trim() === 'term');
  const spendIdx = headers.findIndex(h => String(h).toLowerCase().trim() === 'spend');
  const salesIdx = headers.findIndex(h => String(h).toLowerCase().trim() === 'sales');
  const acosIdx = headers.findIndex(h => String(h).toLowerCase().trim() === 'acos');
  let totalSpend = 0; let totalSales = 0;
  rows.forEach(row => { totalSpend += Number(row.spend || 0); totalSales += Number(row.sales || 0); });
  if (termIdx >= 0) totalRow[termIdx] = 'Total';
  if (spendIdx >= 0) totalRow[spendIdx] = totalSpend;
  if (salesIdx >= 0) totalRow[salesIdx] = totalSales;
  if (acosIdx >= 0) totalRow[acosIdx] = totalSales > 0 ? totalSpend / totalSales : 0;
  const range = sheet.getRange(totalsRowIndex, startCol, 1, headers.length);
  range.setValues([totalRow]); range.setFontWeight('bold');
  if (spendIdx >= 0) sheet.getRange(totalsRowIndex, startCol + spendIdx).setNumberFormat('$#,##0.00');
  if (salesIdx >= 0) sheet.getRange(totalsRowIndex, startCol + salesIdx).setNumberFormat('$#,##0.00');
  if (acosIdx >= 0) sheet.getRange(totalsRowIndex, startCol + acosIdx).setNumberFormat('0.00%');
}
