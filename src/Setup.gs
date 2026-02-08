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
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Rename the first sheet (default "Sheet1") to Contract Pipeline
  renameMainSheet_(ss);

  // Create all required sheets
  ensureSheet_(ss, CONFIG.SHEETS.CONTRACT_PIPELINE);
  ensureSheet_(ss, CONFIG.SHEETS.LEAD_SHEET);
  ensureSheet_(ss, CONFIG.SHEETS.BUYER_SHEET);
  ensureSheet_(ss, CONFIG.SHEETS.KPI_DASHBOARD);
  ensureSheet_(ss, CONFIG.SHEETS.KPI_YEARLY_TRACKER);

  // Set up headers and formatting for each sheet
  setupContractPipelineSheet_(ss);
  setupLeadSheet_(ss);
  setupBuyerSheet_(ss);
  setupKPIDashboardSheet_(ss);
  setupKPIYearlyTrackerSheet_(ss);

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
    '• KPI Dashboard (current month + all time)\n' +
    '• KPI Yearly Tracker (monthly breakdown)\n\n' +
    'KPIs auto-track by month. When a new month starts,\n' +
    'the dashboard automatically shows the new month.\n\n' +
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
    .addItem('View Yearly Tracker', 'goToYearlyTracker')
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
  var sheets = ss.getSheets();
  var first = sheets[0];
  var name = first.getName();

  // Only rename if it's a default name or not already Contract Pipeline
  if (name !== CONFIG.SHEETS.CONTRACT_PIPELINE) {
    // Check if a sheet named "Contract Pipeline" already exists
    var existing = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);
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
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

/**
 * Reorder sheets so Contract Pipeline is the first tab.
 */
function reorderSheets_(ss) {
  var desired = [
    CONFIG.SHEETS.CONTRACT_PIPELINE,
    CONFIG.SHEETS.LEAD_SHEET,
    CONFIG.SHEETS.BUYER_SHEET,
    CONFIG.SHEETS.KPI_DASHBOARD,
    CONFIG.SHEETS.KPI_YEARLY_TRACKER,
  ];

  desired.forEach(function(name, idx) {
    var s = ss.getSheetByName(name);
    if (s) {
      ss.setActiveSheet(s);
      ss.moveActiveSheet(idx + 1);
    }
  });

  // Set Contract Pipeline as the active sheet
  var pipeline = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);
  if (pipeline) ss.setActiveSheet(pipeline);
}

// ── Contract Pipeline Sheet Setup ─────────────────────────────

function setupContractPipelineSheet_(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);
  if (!sheet) return;

  // Clear existing validation and formatting so re-runs work cleanly
  var fullRange = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  fullRange.clearDataValidations();
  sheet.clearConditionalFormatRules();

  var headers = [
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

  // Conditional formatting by Status (must be set BEFORE data validation)
  addStatusConditionalFormatting_(sheet, 16);

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

  // Freeze header row
  sheet.setFrozenRows(1);
}

// ── Lead Sheet Setup ──────────────────────────────────────────

function setupLeadSheet_(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LEAD_SHEET);
  if (!sheet) return;

  // Clear existing validation and formatting so re-runs work cleanly
  var fullRange = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  fullRange.clearDataValidations();
  sheet.clearConditionalFormatRules();

  var headers = [
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
  var sheet = ss.getSheetByName(CONFIG.SHEETS.BUYER_SHEET);
  if (!sheet) return;

  // Clear existing validation and formatting so re-runs work cleanly
  var fullRange = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  fullRange.clearDataValidations();
  sheet.clearConditionalFormatRules();

  var headers = [
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
  var sheet = ss.getSheetByName(CONFIG.SHEETS.KPI_DASHBOARD);
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

  // Last Updated + Current Month
  sheet.getRange('B2').setValue('Last Updated:');
  sheet.getRange('C2').setValue(new Date()).setNumberFormat('MM/dd/yyyy hh:mm AM/PM');
  sheet.getRange('E2').setValue('Current Month:');
  sheet.getRange('F2').setFormula('=TEXT(TODAY(),"MMMM YYYY")');
  sheet.getRange('B2:F2').setFontStyle('italic').setFontColor('#666666');

  // ── Date filter helpers (for formula readability) ──
  // Current month boundaries
  var mStart = 'DATE(YEAR(TODAY()),MONTH(TODAY()),1)';
  var mEnd = 'DATE(YEAR(TODAY()),MONTH(TODAY())+1,1)';

  // ── CURRENT MONTH: Lead & Deal KPIs (Left Column) ──
  var leftMonthly = [
    { label: 'CURRENT MONTH — LEAD & DEAL METRICS', isHeader: true },
    { label: 'New Leads This Month',
      formula: '=COUNTIFS(\'Lead Sheet\'!$A$2:$A,">="&' + mStart + ',\'Lead Sheet\'!$A$2:$A,"<"&' + mEnd + ')' },
    { label: 'Deals Marked Yes',
      formula: '=COUNTIFS(\'Lead Sheet\'!$A$2:$A,">="&' + mStart + ',\'Lead Sheet\'!$A$2:$A,"<"&' + mEnd + ',\'Lead Sheet\'!$S$2:$S,"Yes")' },
    { label: 'Conversion Rate',
      formula: '=IFERROR(COUNTIFS(\'Lead Sheet\'!$A$2:$A,">="&' + mStart + ',\'Lead Sheet\'!$A$2:$A,"<"&' + mEnd + ',\'Lead Sheet\'!$S$2:$S,"Yes")/COUNTIFS(\'Lead Sheet\'!$A$2:$A,">="&' + mStart + ',\'Lead Sheet\'!$A$2:$A,"<"&' + mEnd + '),0)', format: '0.0%' },
    { label: 'Contracts from Lead Sheet',
      formula: '=COUNTIFS(\'Contract Pipeline\'!$B$2:$B,">="&' + mStart + ',\'Contract Pipeline\'!$B$2:$B,"<"&' + mEnd + ',\'Contract Pipeline\'!$A$2:$A,"Lead Sheet")' },
    { label: 'Contracts from Acquisition',
      formula: '=COUNTIFS(\'Contract Pipeline\'!$B$2:$B,">="&' + mStart + ',\'Contract Pipeline\'!$B$2:$B,"<"&' + mEnd + ',\'Contract Pipeline\'!$A$2:$A,"Acquisition")' },
    { label: '', isBlank: true },
    { label: 'CURRENT MONTH — PIPELINE METRICS', isHeader: true },
    { label: 'New Contracts Added',
      formula: '=COUNTIFS(\'Contract Pipeline\'!$B$2:$B,">="&' + mStart + ',\'Contract Pipeline\'!$B$2:$B,"<"&' + mEnd + ')' },
    { label: 'Deals Closed This Month',
      formula: '=COUNTIFS(\'Contract Pipeline\'!$T$2:$T,">="&' + mStart + ',\'Contract Pipeline\'!$T$2:$T,"<"&' + mEnd + ',\'Contract Pipeline\'!$P$2:$P,"Closed")' },
    { label: 'Dead Deals This Month',
      formula: '=COUNTIFS(\'Contract Pipeline\'!$B$2:$B,">="&' + mStart + ',\'Contract Pipeline\'!$B$2:$B,"<"&' + mEnd + ',\'Contract Pipeline\'!$P$2:$P,"Dead")' },
    { label: 'Active Contracts (Under Contract)',
      formula: '=COUNTIF(\'Contract Pipeline\'!$P$2:$P,"Under Contract")' },
  ];

  // ── CURRENT MONTH: Revenue & Buyer KPIs (Right Column) ──
  var rightMonthly = [
    { label: 'CURRENT MONTH — REVENUE METRICS', isHeader: true },
    { label: 'Assignment Fee Revenue',
      formula: '=SUMIFS(\'Contract Pipeline\'!$L$2:$L,\'Contract Pipeline\'!$T$2:$T,">="&' + mStart + ',\'Contract Pipeline\'!$T$2:$T,"<"&' + mEnd + ',\'Contract Pipeline\'!$P$2:$P,"Closed")', format: '$#,##0' },
    { label: 'Avg Assignment Fee',
      formula: '=IFERROR(AVERAGEIFS(\'Contract Pipeline\'!$L$2:$L,\'Contract Pipeline\'!$T$2:$T,">="&' + mStart + ',\'Contract Pipeline\'!$T$2:$T,"<"&' + mEnd + ',\'Contract Pipeline\'!$P$2:$P,"Closed"),0)', format: '$#,##0' },
    { label: CONFIG.KPIS.PIPELINE_VALUE,
      formula: '=SUMIFS(\'Contract Pipeline\'!$L$2:$L,\'Contract Pipeline\'!$P$2:$P,"<>Closed",\'Contract Pipeline\'!$P$2:$P,"<>Dead")', format: '$#,##0' },
    { label: '', isBlank: true },
    { label: '', isBlank: true },
    { label: '', isBlank: true },
    { label: 'CURRENT MONTH — BUYER METRICS', isHeader: true },
    { label: 'New Buyers Onboarded',
      formula: '=COUNTIFS(\'Buyer Sheet\'!$A$2:$A,">="&' + mStart + ',\'Buyer Sheet\'!$A$2:$A,"<"&' + mEnd + ')' },
    { label: CONFIG.KPIS.ACTIVE_BUYERS,
      formula: '=COUNTIF(\'Buyer Sheet\'!$M$2:$M,"Active")' },
    { label: CONFIG.KPIS.CASH_BUYERS,
      formula: '=COUNTIFS(\'Buyer Sheet\'!$F$2:$F,"Cash",\'Buyer Sheet\'!$M$2:$M,"Active")+COUNTIFS(\'Buyer Sheet\'!$F$2:$F,"Both",\'Buyer Sheet\'!$M$2:$M,"Active")' },
    { label: CONFIG.KPIS.AVG_DAYS_TO_CLOSE,
      formula: '=IFERROR(AVERAGE(\'Buyer Sheet\'!$L$2:$L),0)', format: '0' },
  ];

  var startRow = 4;
  setupKPIColumn_(sheet, leftMonthly, startRow, 2, 3);
  setupKPIColumn_(sheet, rightMonthly, startRow, 5, 6);

  // ── Divider row ──
  var dividerRow = startRow + leftMonthly.length + 1;
  sheet.getRange(dividerRow, 2, 1, 5).merge()
    .setBackground('#e0e0e0')
    .setValue('')
    .setBorder(true, true, true, true, false, false);
  sheet.setRowHeight(dividerRow, 8);

  // ── ALL TIME KPIs ──
  var allTimeStartRow = dividerRow + 2;

  // All-time section title
  sheet.getRange(allTimeStartRow - 1, 2).setValue('ALL TIME TOTALS');
  sheet.getRange(allTimeStartRow - 1, 2, 1, 5).merge()
    .setFontSize(14)
    .setFontWeight('bold')
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(allTimeStartRow - 1, 40);

  var leftAllTime = [
    { label: 'ALL TIME — LEAD & DEAL METRICS', isHeader: true },
    { label: CONFIG.KPIS.TOTAL_LEADS, formula: '=COUNTA(\'Lead Sheet\'!A2:A)' },
    { label: CONFIG.KPIS.DEALS_YES, formula: '=COUNTIF(\'Lead Sheet\'!S2:S,"Yes")' },
    { label: CONFIG.KPIS.CONVERSION_RATE, formula: '=IFERROR(COUNTIF(\'Lead Sheet\'!S2:S,"Yes")/COUNTA(\'Lead Sheet\'!A2:A),0)', format: '0.0%' },
    { label: CONFIG.KPIS.DEALS_FROM_LEADS, formula: '=COUNTIF(\'Contract Pipeline\'!A2:A,"Lead Sheet")' },
    { label: CONFIG.KPIS.DEALS_FROM_ACQUISITION, formula: '=COUNTIF(\'Contract Pipeline\'!A2:A,"Acquisition")' },
    { label: '', isBlank: true },
    { label: 'ALL TIME — PIPELINE METRICS', isHeader: true },
    { label: CONFIG.KPIS.ACTIVE_CONTRACTS, formula: '=COUNTIF(\'Contract Pipeline\'!P2:P,"Under Contract")' },
    { label: CONFIG.KPIS.ASSIGNED_DEALS, formula: '=COUNTIF(\'Contract Pipeline\'!P2:P,"Assigned")' },
    { label: CONFIG.KPIS.CLOSED_DEALS, formula: '=COUNTIF(\'Contract Pipeline\'!P2:P,"Closed")' },
    { label: CONFIG.KPIS.DEAD_DEALS, formula: '=COUNTIF(\'Contract Pipeline\'!P2:P,"Dead")' },
  ];

  var rightAllTime = [
    { label: 'ALL TIME — REVENUE METRICS', isHeader: true },
    { label: CONFIG.KPIS.TOTAL_ASSIGNMENT_FEES, formula: '=SUMIF(\'Contract Pipeline\'!P2:P,"Closed",\'Contract Pipeline\'!L2:L)', format: '$#,##0' },
    { label: CONFIG.KPIS.AVG_ASSIGNMENT_FEE, formula: '=IFERROR(AVERAGEIF(\'Contract Pipeline\'!P2:P,"Closed",\'Contract Pipeline\'!L2:L),0)', format: '$#,##0' },
    { label: CONFIG.KPIS.PIPELINE_VALUE, formula: '=SUMIFS(\'Contract Pipeline\'!L2:L,\'Contract Pipeline\'!P2:P,"<>Closed",\'Contract Pipeline\'!P2:P,"<>Dead")', format: '$#,##0' },
    { label: '', isBlank: true },
    { label: '', isBlank: true },
    { label: '', isBlank: true },
    { label: 'ALL TIME — BUYER METRICS', isHeader: true },
    { label: CONFIG.KPIS.TOTAL_BUYERS, formula: '=COUNTA(\'Buyer Sheet\'!A2:A)' },
    { label: CONFIG.KPIS.ACTIVE_BUYERS, formula: '=COUNTIF(\'Buyer Sheet\'!M2:M,"Active")' },
    { label: CONFIG.KPIS.CASH_BUYERS, formula: '=COUNTIFS(\'Buyer Sheet\'!F2:F,"Cash",\'Buyer Sheet\'!M2:M,"Active")+COUNTIFS(\'Buyer Sheet\'!F2:F,"Both",\'Buyer Sheet\'!M2:M,"Active")' },
    { label: CONFIG.KPIS.AVG_DAYS_TO_CLOSE, formula: '=IFERROR(AVERAGE(\'Buyer Sheet\'!L2:L),0)', format: '0' },
  ];

  setupKPIColumn_(sheet, leftAllTime, allTimeStartRow, 2, 3);
  setupKPIColumn_(sheet, rightAllTime, allTimeStartRow, 5, 6);

  sheet.setFrozenRows(1);
}

// ── KPI Yearly Tracker Sheet Setup ────────────────────────────

function setupKPIYearlyTrackerSheet_(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.KPI_YEARLY_TRACKER);
  if (!sheet) return;

  sheet.clear();

  var monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  // ── Column widths ──
  sheet.setColumnWidth(1, 280);   // KPI Name
  for (var c = 2; c <= 13; c++) sheet.setColumnWidth(c, 100);  // Months
  sheet.setColumnWidth(14, 120);  // YTD

  // ── Row 1: Title ──
  sheet.getRange('A1').setValue('KEYSTONE DISPOSITION WORKSPACE — YEARLY KPI TRACKER');
  sheet.getRange('A1:N1').merge()
    .setFontSize(14)
    .setFontWeight('bold')
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 45);

  // ── Row 2: Year + Last Updated ──
  sheet.getRange('A2').setFormula('="Year: "&YEAR(TODAY())').setFontWeight('bold').setFontSize(12);
  sheet.getRange('L2').setValue('Last Updated:').setFontStyle('italic').setFontColor('#666666');
  sheet.getRange('M2').setValue(new Date()).setNumberFormat('MM/dd/yyyy').setFontStyle('italic').setFontColor('#666666');

  // ── Row 3: Column headers ──
  var headerRow = ['KPI Metric'];
  for (var m = 0; m < 12; m++) headerRow.push(monthNames[m]);
  headerRow.push('YTD');
  sheet.getRange(3, 1, 1, 14).setValues([headerRow]);
  sheet.getRange(3, 1, 1, 14)
    .setFontWeight('bold')
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(3, 35);

  // ── Define KPI rows ──
  // Each KPI has a label and a function that returns a formula given month (1-12) or 'ytd'
  var kpiRows = [
    { label: 'LEAD & DEAL METRICS', isHeader: true },
    { label: 'New Leads Received', formulaFn: monthlyLeadCount_ },
    { label: 'Deals Marked Yes', formulaFn: monthlyDealYesCount_ },
    { label: 'Conversion Rate', formulaFn: monthlyConversionRate_, format: '0.0%' },
    { label: '', isBlank: true },
    { label: 'CONTRACT PIPELINE METRICS', isHeader: true },
    { label: 'New Contracts Added', formulaFn: monthlyNewContracts_ },
    { label: 'From Lead Sheet', formulaFn: monthlyContractsFromLeads_ },
    { label: 'From Acquisition', formulaFn: monthlyContractsFromAcq_ },
    { label: 'Deals Closed', formulaFn: monthlyDealsClosed_ },
    { label: 'Dead Deals', formulaFn: monthlyDeadDeals_ },
    { label: '', isBlank: true },
    { label: 'REVENUE METRICS', isHeader: true },
    { label: 'Assignment Fee Revenue', formulaFn: monthlyRevenue_, format: '$#,##0' },
    { label: 'Avg Assignment Fee', formulaFn: monthlyAvgFee_, format: '$#,##0' },
    { label: '', isBlank: true },
    { label: 'BUYER METRICS', isHeader: true },
    { label: 'New Buyers Onboarded', formulaFn: monthlyNewBuyers_ },
  ];

  // ── Write KPI rows ──
  var dataStartRow = 4;

  kpiRows.forEach(function(kpi, idx) {
    var row = dataStartRow + idx;

    if (kpi.isBlank) return;

    if (kpi.isHeader) {
      // Section header — spans all columns
      sheet.getRange(row, 1, 1, 14).merge()
        .setValue(kpi.label)
        .setFontWeight('bold')
        .setFontSize(11)
        .setFontColor(CONFIG.COLORS.HEADER_TEXT)
        .setBackground(CONFIG.COLORS.KPI_HEADER_BG);
      sheet.setRowHeight(row, 30);
      return;
    }

    // KPI label
    sheet.getRange(row, 1).setValue(kpi.label).setFontSize(10);

    // Monthly formulas (columns 2-13 = months 1-12)
    for (var m = 1; m <= 12; m++) {
      var cell = sheet.getRange(row, m + 1);
      cell.setFormula(kpi.formulaFn(m));
      cell.setHorizontalAlignment('center').setFontSize(10);
      if (kpi.format) cell.setNumberFormat(kpi.format);
    }

    // YTD formula (column 14)
    var ytdCell = sheet.getRange(row, 14);
    ytdCell.setFormula(kpi.formulaFn('ytd'));
    ytdCell.setHorizontalAlignment('center')
      .setFontWeight('bold')
      .setFontSize(10)
      .setBackground('#e8f0fe');
    if (kpi.format) ytdCell.setNumberFormat(kpi.format);

    sheet.setRowHeight(row, 28);
  });

  // ── Highlight current month column ──
  // Use conditional formatting to highlight the current month column
  for (var mi = 1; mi <= 12; mi++) {
    var colRange = sheet.getRange(dataStartRow, mi + 1, kpiRows.length, 1);
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=MONTH(TODAY())=' + mi)
      .setBackground('#fff9c4')
      .setRanges([colRange])
      .build();
    var rules = sheet.getConditionalFormatRules();
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);
  }

  sheet.setFrozenRows(3);
  sheet.setFrozenColumns(1);
}

// ── Yearly Tracker Formula Generators ─────────────────────────
// Each function returns a Google Sheets formula string.
// month = 1-12 for a specific month, or 'ytd' for full year.

function monthDateFilter_(dateRange, month) {
  if (month === 'ytd') {
    return dateRange + ',">="&DATE(YEAR(TODAY()),1,1),' + dateRange + ',"<"&DATE(YEAR(TODAY())+1,1,1)';
  }
  return dateRange + ',">="&DATE(YEAR(TODAY()),' + month + ',1),' + dateRange + ',"<"&DATE(YEAR(TODAY()),' + (month + 1) + ',1)';
}

function monthlyLeadCount_(month) {
  return '=COUNTIFS(' + monthDateFilter_('\'Lead Sheet\'!$A$2:$A', month) + ')';
}

function monthlyDealYesCount_(month) {
  return '=COUNTIFS(' + monthDateFilter_('\'Lead Sheet\'!$A$2:$A', month) + ',\'Lead Sheet\'!$S$2:$S,"Yes")';
}

function monthlyConversionRate_(month) {
  var leads = 'COUNTIFS(' + monthDateFilter_('\'Lead Sheet\'!$A$2:$A', month) + ')';
  var yes = 'COUNTIFS(' + monthDateFilter_('\'Lead Sheet\'!$A$2:$A', month) + ',\'Lead Sheet\'!$S$2:$S,"Yes")';
  return '=IFERROR(' + yes + '/' + leads + ',0)';
}

function monthlyNewContracts_(month) {
  return '=COUNTIFS(' + monthDateFilter_('\'Contract Pipeline\'!$B$2:$B', month) + ')';
}

function monthlyContractsFromLeads_(month) {
  return '=COUNTIFS(' + monthDateFilter_('\'Contract Pipeline\'!$B$2:$B', month) + ',\'Contract Pipeline\'!$A$2:$A,"Lead Sheet")';
}

function monthlyContractsFromAcq_(month) {
  return '=COUNTIFS(' + monthDateFilter_('\'Contract Pipeline\'!$B$2:$B', month) + ',\'Contract Pipeline\'!$A$2:$A,"Acquisition")';
}

function monthlyDealsClosed_(month) {
  return '=COUNTIFS(' + monthDateFilter_('\'Contract Pipeline\'!$T$2:$T', month) + ',\'Contract Pipeline\'!$P$2:$P,"Closed")';
}

function monthlyDeadDeals_(month) {
  return '=COUNTIFS(' + monthDateFilter_('\'Contract Pipeline\'!$B$2:$B', month) + ',\'Contract Pipeline\'!$P$2:$P,"Dead")';
}

function monthlyRevenue_(month) {
  return '=SUMIFS(\'Contract Pipeline\'!$L$2:$L,' + monthDateFilter_('\'Contract Pipeline\'!$T$2:$T', month) + ',\'Contract Pipeline\'!$P$2:$P,"Closed")';
}

function monthlyAvgFee_(month) {
  return '=IFERROR(AVERAGEIFS(\'Contract Pipeline\'!$L$2:$L,' + monthDateFilter_('\'Contract Pipeline\'!$T$2:$T', month) + ',\'Contract Pipeline\'!$P$2:$P,"Closed"),0)';
}

function monthlyNewBuyers_(month) {
  return '=COUNTIFS(' + monthDateFilter_('\'Buyer Sheet\'!$A$2:$A', month) + ')';
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
