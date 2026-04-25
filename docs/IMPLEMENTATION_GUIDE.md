# Hybrid Solar + Wind SaaS (Google Sheets + Apps Script)

## 1) Mandatory Sheet Structure

Create these sheets exactly:

1. `INPUT`
   - Column A = field label
   - Column B = value
   - Required labels:
     - Client Name
     - Client Email
     - Location
     - Project Code
     - Solar Capacity (kW)
     - No. of Wind Turbines
     - Wind Turbine Capacity (kW each)
     - Avg Solar Hours/Day
     - Avg Wind Capacity Factor (%)
     - Tariff (₹/kWh)
     - Monthly Consumption (kWh)
     - Total Project Cost (₹)
     - Subsidy / Incentive (₹)
     - Accelerated Depreciation Applicable (YES/NO)
     - Proposal Notes

2. `RESULT`
   - Calculated KPIs written by script.

3. `MONTHLY`
   - Month-wise consumption, generation, savings, and coverage.

4. `CLIENTS`
   - CRM table with status tracking.

5. `CONFIG` (recommended for production controls)
   - Runtime rates such as GST, performance ratios, multipliers.

6. `LOGS`
   - Timestamped logs for diagnostics.

## 2) Core Functions

- `initializeSystem()`
  - Builds/repairs all sheets and seeds config defaults.

- `runFullCalculation()`
  - Reads/validates input.
  - Executes hybrid generation + finance engine.
  - Writes `RESULT` + `MONTHLY`.
  - Upserts client into CRM.

- `generateProposalPdfFromInput()`
  - Builds HTML proposal and exports to PDF in Drive.

- `sendProposalEmailFromInput()`
  - Sends professional proposal email with PDF attachment.
  - Marks CRM status as `Proposal Sent`.

- `upsertClientFromInput()`
  - Creates/updates CRM record by Project Code.

## 3) Financial Model Included

- Solar generation model using PR factor
- Wind generation model using capacity factor
- Hybrid annual generation and annual savings
- GST split model at 5% and 18%
- Net project cost after subsidy
- Depreciation tax shield (accelerated depreciation approximation)
- ROI and payback computation

## 4) Scalability Path

- Web app endpoint (`doGet`) in place for phased productization.
- Config sheet ensures no hardcoded market rates.
- Clean modular functions allow migration to:
  - Multi-user tenancy
  - Payment/subscription layer
  - CRM APIs and WhatsApp automation
  - Data warehouse integration

## 5) Deployment Steps

1. Open Google Sheet.
2. Go to **Extensions → Apps Script**.
3. Paste `apps_script/Code.gs` into `Code.gs`.
4. Add new HTML file named `ProposalTemplate` and paste `apps_script/ProposalTemplate.html`.
5. Save project.
6. Reload spreadsheet.
7. Use menu **Hybrid SaaS → Initialize / Repair Sheets**.
8. Fill `INPUT` values.
9. Run **Hybrid SaaS → Run Full Calculation**.
10. Run **Generate PDF Proposal** and **Send Proposal Email** as needed.

## 6) Production Recommendations

- Use installable triggers for scheduled follow-ups.
- Integrate with Google Cloud Logging via Apps Script advanced services.
- Add row-level ACL model if multiple sales users share one master sheet.
- Store generated PDFs in client/project folder hierarchy.
- Add retry queue for outbound email and webhook actions.
