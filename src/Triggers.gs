/**
 * ============================================================
 * Keystone Disposition Workspace - Trigger Management
 * ============================================================
 *
 * Manages all automated triggers:
 *
 * 1. Form Submit Triggers
 *    - onLeadFormSubmit: When deal form is submitted, check if
 *      Deal Option = "Yes" and push to Contract Pipeline
 *    - onBuyerFormSubmit: When buyer form is submitted, refresh KPIs
 *
 * 2. Time-Driven Triggers
 *    - syncContractPipeline: Runs every N minutes to pull from
 *      Acquisition Workspace and push Deal-Yes leads
 *    - refreshKPIDashboard: Runs after every sync
 *
 * 3. Edit Trigger
 *    - onEditTrigger: Watches for manual status changes in Lead Sheet
 *      (e.g., changing Deal Option from No → Yes) and syncs
 */

/**
 * Install all triggers. Safe to call multiple times — removes
 * existing Keystone triggers before reinstalling.
 */
function installTriggers() {
  removeExistingTriggers_();

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Form submit trigger (fires for any form linked to this spreadsheet)
  ScriptApp.newTrigger('onFormSubmitRouter')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  // 2. Time-driven trigger: sync pipeline every N minutes
  ScriptApp.newTrigger('syncContractPipeline')
    .timeBased()
    .everyMinutes(CONFIG.SYNC_INTERVAL_MINUTES)
    .create();

  // 3. On-edit trigger: watch for Deal Option changes
  ScriptApp.newTrigger('onEditTrigger')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  Logger.log('installTriggers: All triggers installed successfully.');
  Logger.log('  - Form submit → onFormSubmitRouter');
  Logger.log('  - Time-driven → syncContractPipeline (every ' + CONFIG.SYNC_INTERVAL_MINUTES + ' min)');
  Logger.log('  - On edit → onEditTrigger');
}

/**
 * Remove all existing project triggers to prevent duplicates.
 */
function removeExistingTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
  });
  Logger.log('removeExistingTriggers_: Removed ' + triggers.length + ' existing trigger(s).');
}

/**
 * List all active triggers (for debugging).
 */
function listActiveTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    Logger.log('No active triggers.');
    return;
  }

  triggers.forEach(function(trigger, idx) {
    Logger.log('Trigger ' + (idx + 1) + ':');
    Logger.log('  Handler: ' + trigger.getHandlerFunction());
    Logger.log('  Event Type: ' + trigger.getEventType());
    Logger.log('  Trigger Source: ' + trigger.getTriggerSource());
  });
}

// ── Trigger Handlers ──────────────────────────────────────────

/**
 * Routes form submissions to the appropriate handler based on
 * which sheet received the form response.
 */
function onFormSubmitRouter(e) {
  if (!e || !e.range) {
    Logger.log('onFormSubmitRouter: No event range. Skipping.');
    return;
  }

  var sheetName = e.range.getSheet().getName();
  Logger.log('onFormSubmitRouter: Form submitted to sheet "' + sheetName + '"');

  if (sheetName === CONFIG.SHEETS.LEAD_SHEET) {
    onLeadFormSubmit(e);
  } else if (sheetName === CONFIG.SHEETS.BUYER_SHEET) {
    onBuyerFormSubmit(e);
  } else {
    Logger.log('onFormSubmitRouter: Unrecognized sheet "' + sheetName + '". No action taken.');
  }
}

/**
 * Fires on any edit in the spreadsheet.
 * Specifically watches for:
 *   - Deal Option changing to "Yes" in Lead Sheet → push to pipeline
 *   - Status changes in Contract Pipeline → refresh KPIs
 */
function onEditTrigger(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  var editedRow = e.range.getRow();
  var editedCol = e.range.getColumn();

  // ── Lead Sheet: Deal Option changed ──
  if (sheetName === CONFIG.SHEETS.LEAD_SHEET) {
    // Deal Option is column index 18 (0-based) = column 19 (1-based)
    var dealOptionCol = CONFIG.LEAD_COLUMNS.DEAL_OPTION + 1;

    if (editedCol === dealOptionCol && editedRow > 1) {
      var newValue = String(e.range.getValue()).trim();
      Logger.log('onEditTrigger: Deal Option changed to "' + newValue + '" in row ' + editedRow);

      if (newValue.toLowerCase() === 'yes') {
        // Get the full row and push to pipeline
        var row = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
        var address = String(row[CONFIG.LEAD_COLUMNS.PROPERTY_ADDRESS]).trim();

        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var pipelineSheet = ss.getSheetByName(CONFIG.SHEETS.CONTRACT_PIPELINE);
        if (!pipelineSheet) return;

        var existingAddresses = getExistingPipelineAddresses_(pipelineSheet);
        if (address && !existingAddresses[address.toLowerCase()]) {
          var pipelineRow = buildPipelineRowFromLead_(row);
          appendRowsToPipeline_(pipelineSheet, [pipelineRow]);
          refreshKPIDashboard();
          Logger.log('onEditTrigger: Added "' + address + '" to Contract Pipeline.');
        }
      }
    }
  }

  // ── Contract Pipeline: Status changed ──
  if (sheetName === CONFIG.SHEETS.CONTRACT_PIPELINE) {
    var statusCol = CONFIG.CONTRACT_PIPELINE_COLUMNS.STATUS + 1;
    if (editedCol === statusCol && editedRow > 1) {
      refreshKPIDashboard();
      Logger.log('onEditTrigger: Pipeline status changed. KPIs refreshed.');
    }
  }
}
