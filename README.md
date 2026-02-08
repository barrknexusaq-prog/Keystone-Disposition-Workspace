# Keystone Disposition Workspace

Google Apps Script automation for the Keystone Disposition Google Sheets workspace. Manages the full disposition pipeline from lead intake through buyer assignment and closing.

## Architecture

```
┌─────────────────────┐     ┌─────────────────────┐
│  Deal/Partnership    │     │  Buyer Onboarding    │
│  Google Form         │     │  Google Form         │
└────────┬────────────┘     └────────┬────────────┘
         │ auto                      │ auto
         ▼                           ▼
┌─────────────────────┐     ┌─────────────────────┐
│  Lead Sheet          │     │  Buyer Sheet         │
│  (all submissions)   │     │  (all buyers)        │
└────────┬────────────┘     └──────────────────────┘
         │                           │
         │ Deal Option = "Yes"       │ Buyer metrics
         ▼                           ▼
┌──────────────────────────────────────────────────┐
│             CONTRACT PIPELINE (Main Page)         │
│  ← Also receives from Acquisition Workspace      │
│     where Status = "Under Contract"               │
└────────────────────┬─────────────────────────────┘
                     │
                     ▼
┌──────────────────────────────────────────────────┐
│              KPI DASHBOARD (Auto-calculated)      │
│  Lead metrics, pipeline metrics, revenue, buyers  │
└──────────────────────────────────────────────────┘
```

## Sheets

| Sheet | Purpose |
|-------|---------|
| **Contract Pipeline** | Main page. Deals under contract from both Lead Sheet and Acquisition Workspace. |
| **Lead Sheet** | Auto-populated by the deal/partnership Google Form. |
| **Buyer Sheet** | Auto-populated by the buyer onboarding Google Form. |
| **KPI Dashboard** | Auto-calculated metrics across all sheets. |

## Automated Data Flow

1. **Lead Sheet → Contract Pipeline**: Any row where `Deal Option = "Yes"` is automatically copied to the Contract Pipeline with `Source = "Lead Sheet"`. Triggers on form submit, manual edit, and scheduled sync.

2. **Acquisition Workspace → Contract Pipeline**: Any row in the Acquisition workspace's Contract Pipeline sheet where `Status = "Under Contract"` is automatically pulled into this Dispo workspace with `Source = "Acquisition"`. Runs on a scheduled sync (default: every 5 minutes).

3. **Deduplication**: Property Address is used as the unique key. A deal will not be added twice.

4. **KPI Dashboard**: Uses live spreadsheet formulas that recalculate automatically. Timestamp updates on every sync.

## KPIs Tracked

### Lead & Deal Metrics
- Total Leads Received
- Deals Marked Yes
- Lead-to-Deal Conversion Rate
- Contracts from Lead Sheet
- Contracts from Acquisition

### Contract Pipeline Metrics
- Active Contracts (Under Contract)
- Deals Assigned to Buyers
- Deals Closed
- Dead Deals

### Revenue Metrics
- Total Assignment Fee Revenue
- Average Assignment Fee
- Pipeline Value (Pending Assignment Fees)

### Buyer Metrics
- Total Buyers
- Active Buyers
- Cash Buyers
- Avg Days to Close

## Deployment

### Option A: Copy-paste into Apps Script editor

1. Open your Google Spreadsheet
2. Go to **Extensions → Apps Script**
3. Delete any existing code in `Code.gs`
4. Create the following files and paste the corresponding code:
   - `Config.gs` ← `src/Config.gs`
   - `Setup.gs` ← `src/Setup.gs`
   - `ContractPipeline.gs` ← `src/ContractPipeline.gs`
   - `KPIDashboard.gs` ← `src/KPIDashboard.gs`
   - `Triggers.gs` ← `src/Triggers.gs`
5. In `Config.gs`, update `ACQUISITION_WORKSPACE_ID` with your Acquisition spreadsheet ID
6. Save all files (Ctrl+S)
7. Run `setupWorkspace` from the editor (select it from the function dropdown and click Run)
8. Authorize the script when prompted

### Option B: Deploy with clasp

1. Install clasp: `npm install -g @google/clasp`
2. Login: `clasp login`
3. Update `.clasp.json` with your Apps Script project ID
4. Push: `clasp push`
5. Open the script editor: `clasp open`
6. Run `setupWorkspace`

## Configuration

Edit `Config.gs` to customize:

- **ACQUISITION_WORKSPACE_ID**: The spreadsheet ID of your Acquisition Google Workspace
- **Column mappings**: Adjust if your Google Form fields are in a different order
- **SYNC_INTERVAL_MINUTES**: How often to pull from Acquisition (default: 5)
- **COLORS**: Header and status row colors

## Custom Menu

After setup, a **Keystone Dispo** menu appears in the spreadsheet:

- **Run Full Setup** — Re-run the complete setup
- **Sync Contract Pipeline Now** — Manual sync from both sources
- **Refresh KPIs Now** — Force-refresh the KPI dashboard
- **Sync from Acquisition Workspace** — Pull only from Acquisition
- **Push Deal-Yes to Pipeline** — Push only from Lead Sheet

## Linking Google Forms

1. Open your Google Form for deals/partnerships
2. Go to **Responses → Link to Sheets** → select this spreadsheet → select "Lead Sheet"
3. Open your Google Form for buyer onboarding
4. Go to **Responses → Link to Sheets** → select this spreadsheet → select "Buyer Sheet"

The form-submit trigger will automatically detect which sheet received the response and route it accordingly.
