/**
 * ============================================================
 * Keystone Disposition Workspace - Contract Pipeline Automation
 * ============================================================
 *
 * Two data sources feed into the Contract Pipeline:
 *
 * 1. Lead Sheet (this spreadsheet)
 *    - Rows where "Deal Option" = "Yes" are auto-copied
 *
 * 2. Acquisition Workspace (external spreadsheet)
 *    - Rows where "Status" = "Under Contract" are auto-copied
 *
 * Deduplication is based on Property Address to prevent
 * duplicate entries when syncs run repeatedly.
 */

/**
 * Master sync function — runs both sources.
 * Called by time-driven trigger and manual menu item.
 */
function syncContractPipeline() {
  pushDealYesToPipeline();
  pullFromAcquisitionWorkspace();
  refreshKPIDashboard();
}

// ── Source 1: Lead Sheet → Contract Pipeline ──────────────────

/**
 * Scan the Lead Sheet for rows where Deal Option = "Yes"
 * and copy them to the Contract Pipeline if not already present.
 */
function pushDealYesToPipeline() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadSheet = ss.getSheetByName(CONFIG.SHEETS.LEAD_SHEET);
  var pipelineSheet = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);

  if (!leadSheet || !pipelineSheet) {
    Logger.log('pushDealYesToPipeline: Lead Sheet or Contract Pipeline not found.');
    return;
  }

  var leadData = leadSheet.getDataRange().getValues();
  if (leadData.length <= 1) return; // Only header row

  // Get existing pipeline addresses for deduplication
  var existingAddresses = getExistingPipelineAddresses_(pipelineSheet);

  var cols = CONFIG.LEAD_COLUMNS;
  var newRows = [];

  for (var i = 1; i < leadData.length; i++) {
    var row = leadData[i];
    var dealOption = String(row[cols.DEAL_OPTION]).trim();
    var address = String(row[cols.PROPERTY_ADDRESS]).trim();

    // Skip if not "Yes" or already in pipeline
    if (dealOption.toLowerCase() !== 'yes') continue;
    if (!address || existingAddresses[address.toLowerCase()]) continue;

    // Map Lead Sheet columns → Contract Pipeline columns
    var pipelineRow = buildPipelineRowFromLead_(row);
    newRows.push(pipelineRow);
    existingAddresses[address.toLowerCase()] = true;
  }

  if (newRows.length > 0) {
    appendRowsToPipeline_(pipelineSheet, newRows);
    Logger.log('pushDealYesToPipeline: Added ' + newRows.length + ' deal(s) to Contract Pipeline.');
  } else {
    Logger.log('pushDealYesToPipeline: No new deals to add.');
  }
}

/**
 * Map a Lead Sheet row to a Contract Pipeline row.
 */
function buildPipelineRowFromLead_(leadRow) {
  var cols = CONFIG.LEAD_COLUMNS;
  var pCols = CONFIG.CONTRACT_PIPELINE_COLUMNS;

  var pipelineRow = [];
  for (var x = 0; x < 23; x++) pipelineRow.push('');

  pipelineRow[pCols.SOURCE] = 'Lead Sheet';
  pipelineRow[pCols.DATE_ADDED] = new Date();
  pipelineRow[pCols.PROPERTY_ADDRESS] = leadRow[cols.PROPERTY_ADDRESS];
  pipelineRow[pCols.CITY] = leadRow[cols.CITY];
  pipelineRow[pCols.STATE] = leadRow[cols.STATE];
  pipelineRow[pCols.ZIP] = leadRow[cols.ZIP];
  pipelineRow[pCols.COUNTY] = leadRow[cols.COUNTY];
  pipelineRow[pCols.PROPERTY_TYPE] = leadRow[cols.PROPERTY_TYPE];
  pipelineRow[pCols.ARV] = leadRow[cols.ARV];
  pipelineRow[pCols.REPAIR_COST] = leadRow[cols.REPAIR_COST];
  pipelineRow[pCols.CONTRACT_PRICE] = leadRow[cols.ASKING_PRICE];
  pipelineRow[pCols.STATUS] = 'Under Contract';
  pipelineRow[pCols.NOTES] = leadRow[cols.NOTES] || '';

  return pipelineRow;
}

// ── Source 2: Acquisition Workspace → Contract Pipeline ───────

/**
 * Pull rows from the Acquisition Workspace's Contract Pipeline
 * where Status = "Under Contract" into this Dispo workspace.
 */
function pullFromAcquisitionWorkspace() {
  var acqId = CONFIG.ACQUISITION_WORKSPACE_ID;

  // Guard: skip if Acquisition ID is not configured
  if (!acqId || acqId === 'PASTE_ACQUISITION_SPREADSHEET_ID_HERE') {
    Logger.log('pullFromAcquisitionWorkspace: Acquisition Workspace ID not configured. Skipping.');
    return;
  }

  var acqSS;
  try {
    acqSS = SpreadsheetApp.openById(acqId);
  } catch (e) {
    Logger.log('pullFromAcquisitionWorkspace: Cannot open Acquisition Workspace. ' +
               'Check the ID and sharing permissions. Error: ' + e.message);
    return;
  }

  var acqSheet = acqSS.getSheetByName(CONFIG.ACQUISITION_SHEETS.CONTRACT_PIPELINE);
  if (!acqSheet) {
    Logger.log('pullFromAcquisitionWorkspace: Sheet "' +
               CONFIG.ACQUISITION_SHEETS.CONTRACT_PIPELINE + '" not found in Acquisition Workspace.');
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pipelineSheet = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);
  if (!pipelineSheet) return;

  var acqData = acqSheet.getDataRange().getValues();
  if (acqData.length <= 1) return;

  var existingAddresses = getExistingPipelineAddresses_(pipelineSheet);
  var aCols = CONFIG.ACQUISITION_COLUMNS;
  var newRows = [];

  for (var i = 1; i < acqData.length; i++) {
    var row = acqData[i];
    var status = String(row[aCols.STATUS]).trim();
    var address = String(row[aCols.PROPERTY_ADDRESS]).trim();

    // Only pull "Under Contract" rows
    if (status.toLowerCase() !== 'under contract') continue;
    if (!address || existingAddresses[address.toLowerCase()]) continue;

    var pipelineRow = buildPipelineRowFromAcquisition_(row);
    newRows.push(pipelineRow);
    existingAddresses[address.toLowerCase()] = true;
  }

  if (newRows.length > 0) {
    appendRowsToPipeline_(pipelineSheet, newRows);
    Logger.log('pullFromAcquisitionWorkspace: Added ' + newRows.length + ' deal(s) from Acquisition.');
  } else {
    Logger.log('pullFromAcquisitionWorkspace: No new deals from Acquisition.');
  }
}

/**
 * Map an Acquisition Workspace row to a Contract Pipeline row.
 */
function buildPipelineRowFromAcquisition_(acqRow) {
  var aCols = CONFIG.ACQUISITION_COLUMNS;
  var pCols = CONFIG.CONTRACT_PIPELINE_COLUMNS;

  var pipelineRow = [];
  for (var x = 0; x < 23; x++) pipelineRow.push('');

  pipelineRow[pCols.SOURCE] = 'Acquisition';
  pipelineRow[pCols.DATE_ADDED] = new Date();
  pipelineRow[pCols.PROPERTY_ADDRESS] = acqRow[aCols.PROPERTY_ADDRESS];
  pipelineRow[pCols.CITY] = acqRow[aCols.CITY];
  pipelineRow[pCols.STATE] = acqRow[aCols.STATE];
  pipelineRow[pCols.ZIP] = acqRow[aCols.ZIP];
  pipelineRow[pCols.COUNTY] = acqRow[aCols.COUNTY];
  pipelineRow[pCols.PROPERTY_TYPE] = acqRow[aCols.PROPERTY_TYPE];
  pipelineRow[pCols.ARV] = acqRow[aCols.ARV];
  pipelineRow[pCols.REPAIR_COST] = acqRow[aCols.REPAIR_COST];
  pipelineRow[pCols.CONTRACT_PRICE] = acqRow[aCols.CONTRACT_PRICE];
  pipelineRow[pCols.STATUS] = 'Under Contract';
  pipelineRow[pCols.CLOSING_DATE] = acqRow[aCols.CLOSING_DATE] || '';
  pipelineRow[pCols.NOTES] = acqRow[aCols.NOTES] || '';

  return pipelineRow;
}

// ── Shared Pipeline Helpers ───────────────────────────────────

/**
 * Build a lookup object of existing property addresses in the pipeline.
 * Used for deduplication.
 */
function getExistingPipelineAddresses_(pipelineSheet) {
  var pCols = CONFIG.CONTRACT_PIPELINE_COLUMNS;
  var data = pipelineSheet.getDataRange().getValues();
  var addresses = {};

  for (var i = 1; i < data.length; i++) {
    var addr = String(data[i][pCols.PROPERTY_ADDRESS]).trim().toLowerCase();
    if (addr) addresses[addr] = true;
  }

  return addresses;
}

/**
 * Append new rows to the Contract Pipeline sheet.
 */
function appendRowsToPipeline_(pipelineSheet, rows) {
  var lastRow = pipelineSheet.getLastRow();
  var startRow = lastRow + 1;

  pipelineSheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

// ── Form Submit Handler ───────────────────────────────────────

/**
 * Triggered when a form response is submitted to the Lead Sheet.
 * Checks if the new submission has Deal Option = "Yes" and
 * immediately adds it to the Contract Pipeline.
 */
function onLeadFormSubmit(e) {
  if (!e || !e.range) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = e.range.getSheet();

  // Only process if the submission went to the Lead Sheet
  if (sheet.getName() !== CONFIG.SHEETS.LEAD_SHEET) return;

  var row = sheet.getRange(e.range.getRow(), 1, 1, sheet.getLastColumn()).getValues()[0];
  var dealOption = String(row[CONFIG.LEAD_COLUMNS.DEAL_OPTION]).trim();

  if (dealOption.toLowerCase() !== 'yes') return;

  var pipelineSheet = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);
  if (!pipelineSheet) return;

  var address = String(row[CONFIG.LEAD_COLUMNS.PROPERTY_ADDRESS]).trim();
  var existingAddresses = getExistingPipelineAddresses_(pipelineSheet);

  if (!address || existingAddresses[address.toLowerCase()]) return;

  var pipelineRow = buildPipelineRowFromLead_(row);
  appendRowsToPipeline_(pipelineSheet, [pipelineRow]);
  refreshKPIDashboard();

  Logger.log('onLeadFormSubmit: Added "' + address + '" to Contract Pipeline.');
}

/**
 * Triggered when a form response is submitted to the Buyer Sheet.
 * Refreshes KPIs to reflect the new buyer.
 */
function onBuyerFormSubmit(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.SHEETS.BUYER_SHEET) return;

  refreshKPIDashboard();
  Logger.log('onBuyerFormSubmit: Refreshed KPIs after new buyer submission.');
}
