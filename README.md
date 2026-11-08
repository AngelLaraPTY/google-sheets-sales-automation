# Automated Sales Reporting — Google Sheets + Apps Script

A self-contained portfolio demo using fictional sales data. No external API, paid service, or real customer information is required.

## Setup (10–15 minutes)
1. Create a **new blank Google Sheet** named `Automated Sales Reporting — Demo`.
2. Go to **Extensions → Apps Script**. Replace the content of `Code.gs` with the supplied `Code.gs` file, then save.
3. In **Project Settings**, set the script timezone to your preferred timezone (e.g., America/Panama).
4. Reload the Google Sheet. You should see **Sales Automation** in the menu.
5. Click **Sales Automation → Load sample data (demo only)**. Grant the requested Google permissions after reviewing them. This action writes fictional sample data into `Raw_Data` and generates the other tabs.
6. Inspect `Clean_Data`, `Summary`, and `Dashboard`. Try **Sales Automation → Refresh reports**.
7. Optional: **Enable daily refresh**. This creates a time-based trigger, which you can later disable using the same menu. Triggers execute at approximate times, not guaranteed exact times.

## Data contract
`Raw_Data` requires these headers (order may vary):
`Order ID`, `Date`, `Product`, `Category`, `Quantity`, `Unit Price`, `Status`.
- Dates: real Google Sheets dates or text `YYYY-MM-DD`.
- Quantity: positive integer; Unit Price: nonnegative number.
- Only rows with `Status=Completed` are included in revenue.
- Duplicate Order IDs: first valid occurrence wins. Later duplicates are counted and excluded.
- Blank/invalid fields: excluded and counted. Cancelled and other statuses: excluded and counted.
- The script does **not** modify `Raw_Data` when refreshing reports; **Load sample data** overwrites that tab after confirmation when it already has rows.
- A header-only `Raw_Data` sheet is valid and produces zero-value KPIs plus labeled empty chart states.

## Acceptance checks
- After loading sample data, `Raw_Data` has 120 orders + 1 header row.
- `Clean_Data` contains 109 completed orders + 1 header row (11 cancelled).
- `Summary` reports `Orders = 109`, `Invalid Rows = 0`, `Duplicate IDs = 0`, `Excluded Status = 11`.
- `Dashboard` displays three KPI cards, a compact data-quality strip, and two charts.
- Refresh reports several times: the Dashboard must continue to show exactly two charts (when completed-order data exists), with no overlap or duplicate charts.
- Remove all completed orders and refresh: KPI values should show zero and each chart area should show a friendly empty-data message.
- Add a completed order with a unique ID, valid date, quantity and price, then refresh: orders and revenue should increase.
- Add a duplicate order ID, then refresh: duplicate count should increase but orders and revenue should not.
- Change one valid quantity to -1: invalid count should increase after refresh.

## Dashboard design
The Dashboard is intentionally limited to 12 columns and 29 visible rows so its primary content fits comfortably in a typical 1366×768 browser viewport at normal zoom. It uses a navy header, white cards on a light canvas, and green metric/chart accents. Gridlines and unused rows/columns are hidden on this tab only.

The dashboard reads chart data from `Summary`; it does not duplicate helper tables in the presentation area. Every refresh removes existing embedded charts before inserting replacements, preventing chart accumulation. When there are no completed orders, labeled empty states replace the charts. The category visualization uses Google Sheets' native pie chart with a doughnut hole, so there are no external dependencies.

## Calculation regression checks
Each refresh now verifies the established business calculations before replacing report output:

- every source row is accounted for as completed, invalid, duplicate, or excluded;
- all report counters are nonnegative integers;
- detail, monthly, and category revenue totals reconcile within a small floating-point tolerance.

If a future code change breaks one of these invariants, refresh stops with a descriptive `Regression check failed` error instead of publishing inconsistent report output.

## Manual test procedure
1. Load the sample data and confirm the acceptance values above.
2. Run **Sales Automation → Refresh reports** and confirm the Dashboard has no gridlines, all three KPI cards are populated from the report, and both charts are contained below the data-quality strip.
3. Run refresh two more times and verify that only one monthly chart and one category chart remain.
4. Add valid, duplicate, invalid, and non-completed records one at a time. Refresh after each edit and verify the corresponding KPI or quality count against `Clean_Data` and `Summary`.
5. Temporarily change all statuses to a value other than `Completed` (or leave only the required header row), refresh, and verify the zero-value KPI/empty-chart state. Restore the data afterward.
6. Confirm the custom menu still supports manual refresh and that enabling/disabling the daily trigger behaves as before.

These checks must be performed in Google Sheets because embedded charts and sheet viewport behavior are Apps Script runtime features. This repository does not include a local emulator for those services.

## Portfolio honesty and safety
Use screenshots only after actually running the demo. Label all sample figures **fictional demo data**. Do not represent the demo as a client project. Never place customer data, API secrets, or personal identifiers in a public repository. Review Apps Script permission prompts and avoid granting broader access than needed.

## Scope
One spreadsheet, Google Apps Script, local transformations, summary, dashboard, and optional scheduled refresh. No external API integrations or monitoring service.
