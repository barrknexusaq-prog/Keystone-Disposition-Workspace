/**
 * ============================================================
 * Keystone Disposition Workspace - KPI Dashboard Automation
 * ============================================================
 *
 * The KPI Dashboard uses live spreadsheet formulas set during
 * setup (see Setup.gs → setupKPIDashboardSheet_). Those formulas
 * auto-recalculate whenever data changes.
 *
 * This file provides:
 *   - refreshKPIDashboard(): Updates the "Last Updated" timestamp
 *     and forces a recalculation of formula-based KPIs.
 *   - getKPISummary(): Returns a snapshot of all KPIs as an object
 *     (useful for logging, alerts, or external integrations).
 *   - Weekly email digest of KPI performance.
 */

/**
 * Refresh the KPI Dashboard.
 * Updates the timestamp and forces SpreadsheetApp to recalculate.
 */
function refreshKPIDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var kpiSheet = ss.getSheetByName(CONFIG.SHEETS.KPI_DASHBOARD);
  if (!kpiSheet) return;

  // Update "Last Updated" timestamp (row 2, col 3)
  kpiSheet.getRange('C2').setValue(new Date()).setNumberFormat('MM/dd/yyyy hh:mm AM/PM');

  // Force recalculation by flushing
  SpreadsheetApp.flush();

  Logger.log('refreshKPIDashboard: Dashboard refreshed at ' + new Date());
}

/**
 * Get a snapshot of all KPIs as a structured object.
 * Reads calculated values from the KPI Dashboard sheet.
 */
function getKPISummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var kpiSheet = ss.getSheetByName(CONFIG.SHEETS.KPI_DASHBOARD);
  if (!kpiSheet) return null;

  // Read left column KPIs (B/C, rows 4-15)
  var leftData = kpiSheet.getRange('B4:C15').getValues();
  // Read right column KPIs (E/F, rows 4-15)
  var rightData = kpiSheet.getRange('E4:F15').getValues();

  var summary = {};

  // Parse both columns
  var allData = leftData.concat(rightData);
  allData.forEach(function(row) {
    var label = String(row[0]).trim();
    var value = row[1];
    if (label && label !== '' && !isSectionHeader_(label)) {
      summary[label] = value;
    }
  });

  return summary;
}

/**
 * Check if a label is a section header (all caps, no numeric value).
 */
function isSectionHeader_(label) {
  return label === label.toUpperCase() && label.indexOf('METRIC') > -1;
}

/**
 * Send a weekly KPI email digest to specified recipients.
 * Attach this to a weekly time-driven trigger if desired.
 */
function sendWeeklyKPIDigest() {
  var summary = getKPISummary();
  if (!summary) {
    Logger.log('sendWeeklyKPIDigest: No KPI data available.');
    return;
  }

  // Build email body
  var body = 'KEYSTONE DISPOSITION WORKSPACE — WEEKLY KPI REPORT\n';
  body += '=' .repeat(55) + '\n';
  body += 'Generated: ' + new Date().toLocaleDateString() + '\n\n';

  for (var label in summary) {
    var value = summary[label];
    if (typeof value === 'number') {
      if (label.indexOf('Rate') > -1) {
        value = (value * 100).toFixed(1) + '%';
      } else if (label.indexOf('Fee') > -1 || label.indexOf('Revenue') > -1 || label.indexOf('Value') > -1) {
        value = '$' + value.toLocaleString();
      }
    }
    body += label + ': ' + value + '\n';
  }

  body += '\n' + '=' .repeat(55);
  body += '\nView full dashboard: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl();

  // Recipients — update this list
  var recipients = PropertiesService.getScriptProperties().getProperty('KPI_EMAIL_RECIPIENTS');
  if (!recipients) {
    Logger.log('sendWeeklyKPIDigest: No recipients configured. Set KPI_EMAIL_RECIPIENTS in Script Properties.');
    return;
  }

  MailApp.sendEmail({
    to: recipients,
    subject: 'Keystone Dispo — Weekly KPI Report (' + new Date().toLocaleDateString() + ')',
    body: body,
  });

  Logger.log('sendWeeklyKPIDigest: Email sent to ' + recipients);
}

/**
 * Recalculate KPIs using script logic (backup method).
 * This is an alternative to formula-based KPIs. Use this if
 * you prefer script-calculated values over sheet formulas.
 */
function recalculateKPIsFromScript() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadSheet = ss.getSheetByName(CONFIG.SHEETS.LEAD_SHEET);
  var buyerSheet = ss.getSheetByName(CONFIG.SHEETS.BUYER_SHEET);
  var pipelineSheet = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);
  var kpiSheet = ss.getSheetByName(CONFIG.SHEETS.KPI_DASHBOARD);

  if (!leadSheet || !buyerSheet || !pipelineSheet || !kpiSheet) return;

  // ── Lead Metrics ──
  var leadData = leadSheet.getDataRange().getValues();
  var totalLeads = Math.max(0, leadData.length - 1);
  var dealsYes = 0;
  for (var i = 1; i < leadData.length; i++) {
    if (String(leadData[i][CONFIG.LEAD_COLUMNS.DEAL_OPTION]).trim().toLowerCase() === 'yes') {
      dealsYes++;
    }
  }
  var conversionRate = totalLeads > 0 ? dealsYes / totalLeads : 0;

  // ── Pipeline Metrics ──
  var pipelineData = pipelineSheet.getDataRange().getValues();
  var statusCounts = { 'under contract': 0, 'assigned': 0, 'closed': 0, 'dead': 0 };
  var sourceCounts = { 'lead sheet': 0, 'acquisition': 0 };
  var totalFees = 0;
  var closedFees = [];

  for (var j = 1; j < pipelineData.length; j++) {
    var row = pipelineData[j];
    var status = String(row[CONFIG.CONTRACT_PIPELINE_COLUMNS.STATUS]).trim().toLowerCase();
    var source = String(row[CONFIG.CONTRACT_PIPELINE_COLUMNS.SOURCE]).trim().toLowerCase();
    var fee = parseFloat(row[CONFIG.CONTRACT_PIPELINE_COLUMNS.ASSIGNMENT_FEE]) || 0;

    if (statusCounts.hasOwnProperty(status)) statusCounts[status]++;
    if (sourceCounts.hasOwnProperty(source)) sourceCounts[source]++;

    if (status === 'closed') {
      closedFees.push(fee);
      totalFees += fee;
    }
  }

  var avgFee = closedFees.length > 0 ? totalFees / closedFees.length : 0;

  // Pipeline value = assignment fees for non-closed, non-dead deals
  var pipelineValue = 0;
  for (var k = 1; k < pipelineData.length; k++) {
    var pStatus = String(pipelineData[k][CONFIG.CONTRACT_PIPELINE_COLUMNS.STATUS]).trim().toLowerCase();
    if (pStatus !== 'closed' && pStatus !== 'dead') {
      pipelineValue += parseFloat(pipelineData[k][CONFIG.CONTRACT_PIPELINE_COLUMNS.ASSIGNMENT_FEE]) || 0;
    }
  }

  // ── Buyer Metrics ──
  var buyerData = buyerSheet.getDataRange().getValues();
  var totalBuyers = Math.max(0, buyerData.length - 1);
  var activeBuyers = 0;
  var cashBuyers = 0;

  for (var b = 1; b < buyerData.length; b++) {
    var bStatus = String(buyerData[b][CONFIG.BUYER_COLUMNS.STATUS]).trim().toLowerCase();
    var bType = String(buyerData[b][CONFIG.BUYER_COLUMNS.BUYER_TYPE]).trim().toLowerCase();

    if (bStatus === 'active') {
      activeBuyers++;
      if (bType === 'cash' || bType === 'both') cashBuyers++;
    }
  }

  Logger.log('KPI Recalculation Complete:');
  Logger.log('  Total Leads: ' + totalLeads);
  Logger.log('  Deals Yes: ' + dealsYes);
  Logger.log('  Conversion Rate: ' + (conversionRate * 100).toFixed(1) + '%');
  Logger.log('  Under Contract: ' + statusCounts['under contract']);
  Logger.log('  Assigned: ' + statusCounts['assigned']);
  Logger.log('  Closed: ' + statusCounts['closed']);
  Logger.log('  Dead: ' + statusCounts['dead']);
  Logger.log('  Total Fees: $' + totalFees);
  Logger.log('  Avg Fee: $' + avgFee.toFixed(0));
  Logger.log('  Pipeline Value: $' + pipelineValue);
  Logger.log('  Total Buyers: ' + totalBuyers);
  Logger.log('  Active Buyers: ' + activeBuyers);
  Logger.log('  Cash Buyers: ' + cashBuyers);
}
