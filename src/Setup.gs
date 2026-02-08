/**
 * ============================================================
 * Keystone Disposition Workspace - Sheet Setup & Initialization
 * ============================================================
 *
 * Run setupWorkspace() once to create all sheets, headers,
 * formatting, data validation, and conditional formatting.
 *
 * Safe to re-run — it will not overwrite existing data.
 */

/**
 * Main setup entry point. Run this from the Apps Script editor
 * or from the custom menu.
 */
function setupWorkspace() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Rename the first sheet (default "Sheet1") to Contract Pipeline
  renameMainSheet_(ss);

  // Create all required sheets
  ensureSheet_(ss, CONFIG.SHEETS.CONTRACT_PIPELINE);
  ensureSheet_(ss, CONFIG.SHEETS.LEAD_SHEET);
  ensureSheet_(ss, CONFIG.SHEETS.BUYER_SHEET);
  ensureSheet_(ss, CONFIG.SHEETS.KPI_DASHBOARD);

  // Set up headers and formatting for each sheet
  setupContractPipelineSheet_(ss);
  setupLeadSheet_(ss);
  setupBuyerSheet_(ss);
  setupKPIDashboardSheet_(ss);

  // Reorder sheets: Contract Pipeline first
  reorderSheets_(ss);

  // Install triggers
  installTriggers();

  SpreadsheetApp.getUi().alert(
    'Workspace Setup Complete',
    'All sheets have been created and configured.\n\n' +
    '• Contract Pipeline (main page)\n' +
    '• Lead Sheet (linked to deal form)\n' +
    '• Buyer Sheet (linked to buyer form)\n' +
    '• KPI Dashboard (auto-calculated)\n\n' +
    'Triggers are now active.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Add a custom menu to the spreadsheet for easy access.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Keystone Dispo')
    .addItem('Run Full Setup', 'setupWorkspace')
    .addSeparator()
    .addItem('Sync Contract Pipeline Now', 'syncContractPipeline')
    .addItem('Refresh KPIs Now', 'refreshKPIDashboard')
    .addSeparator()
    .addItem('Sync from Acquisition Workspace', 'pullFromAcquisitionWorkspace')
    .addItem('Push Deal-Yes to Pipeline', 'pushDealYesToPipeline')
    .addToUi();
}

// ── Internal Helpers ──────────────────────────────────────────

/**
 * Rename the default first sheet to Contract Pipeline.
 */
function renameMainSheet_(ss) {
  const sheets = ss.getSheets();
  const first = sheets[0];
  const name = first.getName();

  // Only rename if it's a default name or not already Contract Pipeline
  if (name !== CONFIG.SHEETS.CONTRACT_PIPELINE) {
    // Check if a sheet named "Contract Pipeline" already exists
    const existing = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);
    if (existing && existing.getSheetId() !== first.getSheetId()) {
      // Already exists as a different sheet; leave the first sheet alone
      return;
    }
    first.setName(CONFIG.SHEETS.CONTRACT_PIPELINE);
  }
}

/**
 * Create a sheet if it doesn't already exist.
 */
function ensureSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

/**
 * Reorder sheets so Contract Pipeline is the first tab.
 */
function reorderSheets_(ss) {
  const desired = [
    CONFIG.SHEETS.CONTRACT_PIPELINE,
    CONFIG.SHEETS.LEAD_SHEET,
    CONFIG.SHEETS.BUYER_SHEET,
    CONFIG.SHEETS.KPI_DASHBOARD,
  ];

  desired.forEach(function(name, idx) {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(idx + 1);
    }
  });

  // Set Contract Pipeline as the active sheet
  const pipeline = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);
  if (pipeline) ss.setActiveSheet(pipeline);
}

// ── Contract Pipeline Sheet Setup ─────────────────────────────

function setupContractPipelineSheet_(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);
  if (!sheet) return;

  const headers = [
    'Source',
    'Date Added',
    'Property Address',
    'City',
    'State',
    'Zip',
    'County',
    'Property Type',
    'ARV',
    'Repair Cost',
    'Contract Price',
    'Assignment Fee',
    'Buyer Assigned',
    'Buyer Phone',
    'Buyer Email',
    'Status',
    'Earnest Money',
    'EMD Status',
    'Inspection Period End',
    'Closing Date',
    'Title Company',
    'Disposition Manager',
    'Notes',
  ];

  setHeaders_(sheet, headers);
  formatHeaderRow_(sheet, headers.length);

  // Column widths
  sheet.setColumnWidth(1, 120);   // Source
  sheet.setColumnWidth(2, 110);   // Date Added
  sheet.setColumnWidth(3, 220);   // Property Address
  sheet.setColumnWidth(4, 120);   // City
  sheet.setColumnWidth(5, 60);    // State
  sheet.setColumnWidth(6, 80);    // Zip
  sheet.setColumnWidth(7, 120);   // County
  sheet.setColumnWidth(8, 120);   // Property Type
  sheet.setColumnWidth(9, 110);   // ARV
  sheet.setColumnWidth(10, 110);  // Repair Cost
  sheet.setColumnWidth(11, 120);  // Contract Price
  sheet.setColumnWidth(12, 120);  // Assignment Fee
  sheet.setColumnWidth(13, 150);  // Buyer Assigned
  sheet.setColumnWidth(14, 130);  // Buyer Phone
  sheet.setColumnWidth(15, 180);  // Buyer Email
  sheet.setColumnWidth(16, 130);  // Status
  sheet.setColumnWidth(17, 120);  // Earnest Money
  sheet.setColumnWidth(18, 110);  // EMD Status
  sheet.setColumnWidth(19, 150);  // Inspection End
  sheet.setColumnWidth(20, 120);  // Closing Date
  sheet.setColumnWidth(21, 150);  // Title Company
  sheet.setColumnWidth(22, 150);  // Disposition Manager
  sheet.setColumnWidth(23, 250);  // Notes

  // Data validation: Status column
  var statusRange = sheet.getRange(2, 16, sheet.getMaxRows() - 1, 1);
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Under Contract', 'Assigned', 'Closed', 'Dead'], true)
    .setAllowInvalid(false)
    .build();
  statusRange.setDataValidation(statusRule);

  // Data validation: EMD Status column
  var emdRange = sheet.getRange(2, 18, sheet.getMaxRows() - 1, 1);
  var emdRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Deposited', 'Refunded'], true)
    .setAllowInvalid(false)
    .build();
  emdRange.setDataValidation(emdRule);

  // Data validation: Source column
  var sourceRange = sheet.getRange(2, 1, sheet.getMaxRows() - 1, 1);
  var sourceRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Lead Sheet', 'Acquisition'], true)
    .setAllowInvalid(false)
    .build();
  sourceRange.setDataValidation(sourceRule);

  // Currency formatting for money columns (ARV, Repair, Contract Price, Assignment Fee, Earnest Money)
  var moneyColumns = [9, 10, 11, 12, 17];
  moneyColumns.forEach(function(col) {
    sheet.getRange(2, col, sheet.getMaxRows() - 1, 1).setNumberFormat('$#,##0');
  });

  // Date formatting
  var dateColumns = [2, 19, 20];
  dateColumns.forEach(function(col) {
    sheet.getRange(2, col, sheet.getMaxRows() - 1, 1).setNumberFormat('MM/dd/yyyy');
  });

  // Conditional formatting by Status
  addStatusConditionalFormatting_(sheet, 16);

  // Freeze header row
  sheet.setFrozenRows(1);
}

// ── Lead Sheet Setup ──────────────────────────────────────────

function setupLeadSheet_(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.LEAD_SHEET);
  if (!sheet) return;

  const headers = [
    'Timestamp',
    'Property Address',
    'City',
    'State',
    'Zip',
    'County',
    'Property Type',
    'Bedrooms',
    'Bathrooms',
    'Sq Ft',
    'Lot Size',
    'Year Built',
    'ARV',
    'Repair Cost',
    'Asking Price',
    'Seller Name',
    'Seller Phone',
    'Seller Email',
    'Deal Option',
    'Lead Source',
    'Notes',
  ];

  setHeaders_(sheet, headers);
  formatHeaderRow_(sheet, headers.length);

  // Column widths
  sheet.setColumnWidth(1, 150);   // Timestamp
  sheet.setColumnWidth(2, 220);   // Property Address
  sheet.setColumnWidth(3, 120);   // City
  sheet.setColumnWidth(4, 60);    // State
  sheet.setColumnWidth(5, 80);    // Zip
  sheet.setColumnWidth(19, 100);  // Deal Option

  // Data validation: Deal Option
  var dealRange = sheet.getRange(2, 19, sheet.getMaxRows() - 1, 1);
  var dealRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true)
    .setAllowInvalid(false)
    .build();
  dealRange.setDataValidation(dealRule);

  // Currency formatting
  var moneyColumns = [13, 14, 15];
  moneyColumns.forEach(function(col) {
    sheet.getRange(2, col, sheet.getMaxRows() - 1, 1).setNumberFormat('$#,##0');
  });

  // Highlight rows where Deal Option = Yes
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Yes')
    .setBackground('#d4edda')
    .setRanges([sheet.getRange(2, 19, sheet.getMaxRows() - 1, 1)])
    .build();
  sheet.setConditionalFormatRules([rule]);

  sheet.setFrozenRows(1);
}

// ── Buyer Sheet Setup ─────────────────────────────────────────

function setupBuyerSheet_(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.BUYER_SHEET);
  if (!sheet) return;

  const headers = [
    'Timestamp',
    'Buyer Name',
    'Company',
    'Phone',
    'Email',
    'Buyer Type',
    'Target Markets',
    'Property Types',
    'Min Price',
    'Max Price',
    'Proof of Funds',
    'Closing Speed (Days)',
    'Status',
    'Notes',
  ];

  setHeaders_(sheet, headers);
  formatHeaderRow_(sheet, headers.length);

  // Column widths
  sheet.setColumnWidth(1, 150);   // Timestamp
  sheet.setColumnWidth(2, 160);   // Buyer Name
  sheet.setColumnWidth(3, 160);   // Company
  sheet.setColumnWidth(4, 130);   // Phone
  sheet.setColumnWidth(5, 200);   // Email
  sheet.setColumnWidth(6, 110);   // Buyer Type
  sheet.setColumnWidth(7, 180);   // Target Markets
  sheet.setColumnWidth(8, 150);   // Property Types

  // Data validation: Buyer Type
  var typeRange = sheet.getRange(2, 6, sheet.getMaxRows() - 1, 1);
  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Cash', 'Financing', 'Both'], true)
    .setAllowInvalid(false)
    .build();
  typeRange.setDataValidation(typeRule);

  // Data validation: Proof of Funds
  var pofRange = sheet.getRange(2, 11, sheet.getMaxRows() - 1, 1);
  var pofRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true)
    .setAllowInvalid(false)
    .build();
  pofRange.setDataValidation(pofRule);

  // Data validation: Status
  var statusRange = sheet.getRange(2, 13, sheet.getMaxRows() - 1, 1);
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive'], true)
    .setAllowInvalid(false)
    .build();
  statusRange.setDataValidation(statusRule);

  // Currency formatting
  sheet.getRange(2, 9, sheet.getMaxRows() - 1, 1).setNumberFormat('$#,##0');
  sheet.getRange(2, 10, sheet.getMaxRows() - 1, 1).setNumberFormat('$#,##0');

  sheet.setFrozenRows(1);
}

// ── KPI Dashboard Sheet Setup ─────────────────────────────────

function setupKPIDashboardSheet_(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.KPI_DASHBOARD);
  if (!sheet) return;

  // Clear any existing content
  sheet.clear();

  // Set column widths
  sheet.setColumnWidth(1, 50);    // Spacer
  sheet.setColumnWidth(2, 300);   // KPI Name
  sheet.setColumnWidth(3, 180);   // KPI Value
  sheet.setColumnWidth(4, 50);    // Spacer
  sheet.setColumnWidth(5, 300);   // KPI Name
  sheet.setColumnWidth(6, 180);   // KPI Value

  // Title
  sheet.getRange('B1').setValue('KEYSTONE DISPOSITION WORKSPACE — KPI DASHBOARD');
  sheet.getRange('B1:F1').merge()
    .setFontSize(16)
    .setFontWeight('bold')
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 50);

  // Last Updated
  sheet.getRange('B2').setValue('Last Updated:');
  sheet.getRange('C2').setValue(new Date()).setNumberFormat('MM/dd/yyyy hh:mm AM/PM');
  sheet.getRange('B2:C2').setFontStyle('italic').setFontColor('#666666');

  // ── Lead & Deal KPIs (Left Column) ──
  var leftKPIs = [
    { label: 'LEAD & DEAL METRICS', isHeader: true },
    { label: CONFIG.KPIS.TOTAL_LEADS, formula: '=COUNTA(\'Lead Sheet\'!A2:A)' },
    { label: CONFIG.KPIS.DEALS_YES, formula: '=COUNTIF(\'Lead Sheet\'!S2:S,"Yes")' },
    { label: CONFIG.KPIS.CONVERSION_RATE, formula: '=IFERROR(COUNTIF(\'Lead Sheet\'!S2:S,"Yes")/COUNTA(\'Lead Sheet\'!A2:A),0)', format: '0.0%' },
    { label: CONFIG.KPIS.DEALS_FROM_LEADS, formula: '=COUNTIF(\'Contract Pipeline\'!A2:A,"Lead Sheet")' },
    { label: CONFIG.KPIS.DEALS_FROM_ACQUISITION, formula: '=COUNTIF(\'Contract Pipeline\'!A2:A,"Acquisition")' },
    { label: '', isBlank: true },
    { label: 'CONTRACT PIPELINE METRICS', isHeader: true },
    { label: CONFIG.KPIS.ACTIVE_CONTRACTS, formula: '=COUNTIF(\'Contract Pipeline\'!P2:P,"Under Contract")' },
    { label: CONFIG.KPIS.ASSIGNED_DEALS, formula: '=COUNTIF(\'Contract Pipeline\'!P2:P,"Assigned")' },
    { label: CONFIG.KPIS.CLOSED_DEALS, formula: '=COUNTIF(\'Contract Pipeline\'!P2:P,"Closed")' },
    { label: CONFIG.KPIS.DEAD_DEALS, formula: '=COUNTIF(\'Contract Pipeline\'!P2:P,"Dead")' },
  ];

  // ── Revenue & Buyer KPIs (Right Column) ──
  var rightKPIs = [
    { label: 'REVENUE METRICS', isHeader: true },
    { label: CONFIG.KPIS.TOTAL_ASSIGNMENT_FEES, formula: '=SUMIF(\'Contract Pipeline\'!P2:P,"Closed",\'Contract Pipeline\'!L2:L)', format: '$#,##0' },
    { label: CONFIG.KPIS.AVG_ASSIGNMENT_FEE, formula: '=IFERROR(AVERAGEIF(\'Contract Pipeline\'!P2:P,"Closed",\'Contract Pipeline\'!L2:L),0)', format: '$#,##0' },
    { label: CONFIG.KPIS.PIPELINE_VALUE, formula: '=SUMIFS(\'Contract Pipeline\'!L2:L,\'Contract Pipeline\'!P2:P,"<>Closed",\'Contract Pipeline\'!P2:P,"<>Dead")', format: '$#,##0' },
    { label: '', isBlank: true },
    { label: '', isBlank: true },
    { label: '', isBlank: true },
    { label: 'BUYER METRICS', isHeader: true },
    { label: CONFIG.KPIS.TOTAL_BUYERS, formula: '=COUNTA(\'Buyer Sheet\'!A2:A)' },
    { label: CONFIG.KPIS.ACTIVE_BUYERS, formula: '=COUNTIF(\'Buyer Sheet\'!M2:M,"Active")' },
    { label: CONFIG.KPIS.CASH_BUYERS, formula: '=COUNTIFS(\'Buyer Sheet\'!F2:F,"Cash",\'Buyer Sheet\'!M2:M,"Active")+COUNTIFS(\'Buyer Sheet\'!F2:F,"Both",\'Buyer Sheet\'!M2:M,"Active")' },
    { label: CONFIG.KPIS.AVG_DAYS_TO_CLOSE, formula: '=IFERROR(AVERAGE(\'Buyer Sheet\'!L2:L),0)', format: '0' },
  ];

  var startRow = 4;
  setupKPIColumn_(sheet, leftKPIs, startRow, 2, 3);
  setupKPIColumn_(sheet, rightKPIs, startRow, 5, 6);

  sheet.setFrozenRows(1);
}

function setupKPIColumn_(sheet, kpis, startRow, labelCol, valueCol) {
  kpis.forEach(function(kpi, idx) {
    var row = startRow + idx;
    if (kpi.isBlank) return;

    if (kpi.isHeader) {
      sheet.getRange(row, labelCol).setValue(kpi.label)
        .setFontWeight('bold')
        .setFontSize(12)
        .setFontColor(CONFIG.COLORS.HEADER_TEXT)
        .setBackground(CONFIG.COLORS.KPI_HEADER_BG);
      sheet.getRange(row, valueCol)
        .setBackground(CONFIG.COLORS.KPI_HEADER_BG);
      sheet.setRowHeight(row, 35);
    } else {
      sheet.getRange(row, labelCol).setValue(kpi.label)
        .setFontSize(11)
        .setVerticalAlignment('middle');
      sheet.getRange(row, valueCol).setFormula(kpi.formula)
        .setFontSize(14)
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setBackground(CONFIG.COLORS.KPI_VALUE_BG)
        .setFontColor(CONFIG.COLORS.HEADER_TEXT);

      if (kpi.format) {
        sheet.getRange(row, valueCol).setNumberFormat(kpi.format);
      }

      sheet.setRowHeight(row, 40);
    }
  });
}

// ── Shared Formatting Helpers ─────────────────────────────────

/**
 * Set header row values if the first row is empty.
 */
function setHeaders_(sheet, headers) {
  var firstCell = sheet.getRange(1, 1).getValue();
  if (firstCell === '' || firstCell === null) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

/**
 * Format the header row with consistent styling.
 */
function formatHeaderRow_(sheet, numCols) {
  var headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange
    .setFontWeight('bold')
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontSize(10)
    .setWrap(true);
  sheet.setRowHeight(1, 40);
}

/**
 * Add conditional formatting to color-code rows by status.
 */
function addStatusConditionalFormatting_(sheet, statusCol) {
  var range = sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns());

  var rules = [];

  // Under Contract = yellow
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + columnLetter_(statusCol) + '2="Under Contract"')
    .setBackground(CONFIG.COLORS.UNDER_CONTRACT)
    .setRanges([range])
    .build());

  // Assigned = blue
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + columnLetter_(statusCol) + '2="Assigned"')
    .setBackground(CONFIG.COLORS.ASSIGNED)
    .setRanges([range])
    .build());

  // Closed = green
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + columnLetter_(statusCol) + '2="Closed"')
    .setBackground(CONFIG.COLORS.CLOSED)
    .setRanges([range])
    .build());

  // Dead = red
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + columnLetter_(statusCol) + '2="Dead"')
    .setBackground(CONFIG.COLORS.DEAD)
    .setRanges([range])
    .build());

  sheet.setConditionalFormatRules(rules);
}

/**
 * Convert column number (1-indexed) to letter (A, B, ... Z, AA, AB, ...).
 */
function columnLetter_(col) {
  var letter = '';
  while (col > 0) {
    var mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}
