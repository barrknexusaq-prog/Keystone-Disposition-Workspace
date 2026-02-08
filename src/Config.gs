/**
 * ============================================================
 * Keystone Disposition Workspace - Configuration
 * ============================================================
 *
 * UPDATE THESE VALUES before running setup.
 *
 * SPREADSHEET_IDS:
 *   - DISPO_WORKSPACE: The ID of THIS spreadsheet (Disposition Workspace)
 *   - ACQUISITION_WORKSPACE: The ID of the Acquisition Google Workspace spreadsheet
 *
 * You can find a spreadsheet ID in the URL:
 *   https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit
 */

var CONFIG = {
  // ── Spreadsheet IDs ───────────────────────────────────────
  DISPO_WORKSPACE_ID: '1swI3l9sb7wKLyDRBrIwAdkscN5sU_4nF2tbCNNeMd90',
  ACQUISITION_WORKSPACE_ID: '1VZl6IvUpFX04LpsiXRGXWPui6hTV7pEqPDctkG0rxg4',

  // ── Sheet Names (Disposition Workspace) ───────────────────
  SHEETS: {
    CONTRACT_PIPELINE: 'Contract Pipeline',
    LEAD_SHEET: 'Lead Sheet',
    BUYER_SHEET: 'Buyer Sheet',
    KPI_DASHBOARD: 'KPI Dashboard',
    KPI_YEARLY_TRACKER: 'KPI Yearly Tracker',
  },

  // ── Sheet Names (Acquisition Workspace) ───────────────────
  ACQUISITION_SHEETS: {
    CONTRACT_PIPELINE: 'Contract Pipeline',
  },

  // ── Column Mappings: Lead Sheet ───────────────────────────
  // These map to your Google Form for deals/partnership submissions.
  // Adjust column letters/indices if your form fields differ.
  LEAD_COLUMNS: {
    TIMESTAMP: 0,
    PROPERTY_ADDRESS: 1,
    CITY: 2,
    STATE: 3,
    ZIP: 4,
    COUNTY: 5,
    PROPERTY_TYPE: 6,
    BEDROOMS: 7,
    BATHROOMS: 8,
    SQ_FT: 9,
    LOT_SIZE: 10,
    YEAR_BUILT: 11,
    ARV: 12,
    REPAIR_COST: 13,
    ASKING_PRICE: 14,
    SELLER_NAME: 15,
    SELLER_PHONE: 16,
    SELLER_EMAIL: 17,
    DEAL_OPTION: 18,       // "Yes" / "No" — key trigger column
    LEAD_SOURCE: 19,
    NOTES: 20,
  },

  // ── Column Mappings: Buyer Sheet ──────────────────────────
  // These map to your Google Form for buyer onboarding.
  BUYER_COLUMNS: {
    TIMESTAMP: 0,
    BUYER_NAME: 1,
    COMPANY: 2,
    PHONE: 3,
    EMAIL: 4,
    BUYER_TYPE: 5,         // Cash / Financing / Both
    MARKETS: 6,            // Target markets/areas
    PROPERTY_TYPES: 7,     // SFR, Multi, Land, etc.
    MIN_PRICE: 8,
    MAX_PRICE: 9,
    PROOF_OF_FUNDS: 10,    // Yes / No
    CLOSING_SPEED: 11,     // Days to close
    STATUS: 12,            // Active / Inactive
    NOTES: 13,
  },

  // ── Column Mappings: Contract Pipeline ────────────────────
  CONTRACT_PIPELINE_COLUMNS: {
    SOURCE: 0,             // "Lead Sheet" or "Acquisition"
    DATE_ADDED: 1,
    PROPERTY_ADDRESS: 2,
    CITY: 3,
    STATE: 4,
    ZIP: 5,
    COUNTY: 6,
    PROPERTY_TYPE: 7,
    ARV: 8,
    REPAIR_COST: 9,
    CONTRACT_PRICE: 10,
    ASSIGNMENT_FEE: 11,
    BUYER_ASSIGNED: 12,
    BUYER_PHONE: 13,
    BUYER_EMAIL: 14,
    STATUS: 15,            // Under Contract / Assigned / Closed / Dead
    EARNEST_MONEY: 16,
    EMD_STATUS: 17,        // Deposited / Pending / Refunded
    INSPECTION_END: 18,
    CLOSING_DATE: 19,
    TITLE_COMPANY: 20,
    DISPOSITION_MANAGER: 21,
    NOTES: 22,
  },

  // ── Column Mappings: Acquisition Workspace Contract Pipeline ──
  // Adjust these to match the actual column layout in your Acquisition sheet
  ACQUISITION_COLUMNS: {
    PROPERTY_ADDRESS: 0,
    CITY: 1,
    STATE: 2,
    ZIP: 3,
    COUNTY: 4,
    PROPERTY_TYPE: 5,
    ARV: 6,
    REPAIR_COST: 7,
    CONTRACT_PRICE: 8,
    STATUS: 9,             // "Under Contract" triggers pull
    SELLER_NAME: 10,
    CLOSING_DATE: 11,
    NOTES: 12,
  },

  // ── KPI Definitions ───────────────────────────────────────
  KPIS: {
    TOTAL_LEADS: 'Total Leads Received',
    DEALS_YES: 'Deals Marked Yes',
    CONVERSION_RATE: 'Lead-to-Deal Conversion Rate',
    ACTIVE_CONTRACTS: 'Active Contracts (Under Contract)',
    ASSIGNED_DEALS: 'Deals Assigned to Buyers',
    CLOSED_DEALS: 'Deals Closed',
    DEAD_DEALS: 'Dead Deals',
    TOTAL_ASSIGNMENT_FEES: 'Total Assignment Fee Revenue',
    AVG_ASSIGNMENT_FEE: 'Average Assignment Fee',
    TOTAL_BUYERS: 'Total Buyers',
    ACTIVE_BUYERS: 'Active Buyers',
    CASH_BUYERS: 'Cash Buyers',
    AVG_DAYS_TO_CLOSE: 'Avg Days to Close',
    PIPELINE_VALUE: 'Pipeline Value (Pending Assignment Fees)',
    DEALS_FROM_LEADS: 'Contracts from Lead Sheet',
    DEALS_FROM_ACQUISITION: 'Contracts from Acquisition',
  },

  // ── Trigger Intervals ─────────────────────────────────────
  SYNC_INTERVAL_MINUTES: 5,  // How often to pull from Acquisition Workspace

  // ── Formatting ────────────────────────────────────────────
  COLORS: {
    HEADER_BG: '#1a1a2e',
    HEADER_TEXT: '#ffffff',
    KPI_HEADER_BG: '#16213e',
    KPI_VALUE_BG: '#0f3460',
    UNDER_CONTRACT: '#fff3cd',
    ASSIGNED: '#cce5ff',
    CLOSED: '#d4edda',
    DEAD: '#f8d7da',
    ALTERNATING_ROW: '#f8f9fa',
  },
};
