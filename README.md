# Budget Review App

Local drag-and-drop tool for turning monthly ABN AMRO and Wise exports into a budget review workbook, with optional Google Sheets sync.

The intended user flow is:

1. Export bank transactions.
2. Open the local web app.
3. Drag the files in.
4. Download the generated workbook or sync selected tabs into Google Sheets.

## Quick Start

### Windows

1. Install Python 3.11 or newer from <https://www.python.org/downloads/>.
2. Download or clone this repository.
3. Double-click `start_windows.bat`.
4. The app opens at <http://127.0.0.1:5057>.

The first run creates a local `.venv` folder and installs dependencies. Later runs are faster.

### Mac / Linux

Open a terminal in this folder:

```sh
chmod +x start_macos_linux.sh
./start_macos_linux.sh
```

## Input Files

Supported files:

- ABN AMRO export: `.xls` or `.xlsx`
- Wise transaction history: `.csv`, `.xls`, or `.xlsx`
- Optional existing workbook template: `.xlsx`

ABN exports must contain:

- `transactiondate`
- `amount`
- `description`

Wise exports must keep Wise's default column names, including:

- `Created on`
- `Finished on`
- `Direction`
- `Source amount (after fees)`
- `Target name`
- `Category`

## Monthly Workflow

1. Export ABN AMRO transactions for the review period.
2. Export Wise transaction history for the review period.
3. Start the app.
4. Drop the ABN file into `ABN AMRO export`.
5. Drop the Wise file into `Wise transaction history`.
6. Optionally drop an existing `.xlsx` workbook as a template.
7. Set:
   - period label, for example `July 2026`
   - start date
   - end date
   - data-through date: the latest transaction date included in the files you are uploading
8. Click `Generate review`.
9. Download the workbook, or sync to Google Sheets if configured.

Recommended budget cycle pattern:

- April 2026: `2026-03-25` to `2026-04-24`
- May 2026: `2026-04-25` to `2026-05-24`
- June 2026: `2026-05-25` to `2026-06-24`

This matches the original 25th-to-24th budget cycle and salary timing.

## Output Workbook

Each generated workbook contains:

- `Settings 2026`: editable budgets and income assumptions
- `Raw Txns 2026`: normalized ABN/Wise transactions with category and include/exclude columns
- `Review <Month>`: budget vs actual summary, income/credits grouped by source category, and expenses grouped by budget category

The review tab is formula-driven:

- Budgets link from `Settings 2026`
- Actuals use `SUMIFS` against `Raw Txns 2026`
- Income uses salary and credit/refund rows from `Raw Txns 2026`
- Category variance/status recalculates automatically
- Transaction detail blocks have per-category subtotal formulas

`Actual cash left after expenses` is total income/credits minus actual expenses.

`Actual cash vs budget plan` compares that actual cash left with the cash that would have been left if expenses had exactly matched the monthly budget. Positive means spending came in under plan; negative means expenses were over plan.

## Budget Rules

Current Fun budget split:

- `Ian - Fun`: EUR 250/month
- `Leila - Fun`: EUR 250/month

Current income assumption:

- `Monthly Salary (Leila)`: EUR 0/month

Rules:

- Wise fun spend is assigned to `Ian - Fun`
- Wise `Shopping`, `Personal care`, eating out, entertainment, trips, cash, and general card spend are treated as `Ian - Fun`
- Leila fun-money transfers are assigned to `Leila - Fun`
- ABN-to-Wise top-ups and Wise-to-ABN repayments are `Internal Transfer` and excluded from expenses to avoid double-counting
- Credits/refunds are shown in the income/credits summary, not in expense totals

## Editing After Generation

To adjust a review after generation:

- Change monthly budgets in `Settings 2026`
- Change Leila salary assumption in `Settings 2026`
- Change transaction categories in `Raw Txns 2026`
- Change `Include` from `1` to `0` in `Raw Txns 2026` to exclude a transaction
- Change `Include` from `0` to `1` to include a transaction

Avoid typing over formulas on `Review <Month>` unless you intentionally want to break the live link.

## Google Sheets Sync

Downloading the workbook works immediately.

Direct Google Sheets sync needs a one-time Apps Script setup because a local app cannot edit Google Drive files without permission.

### One-Time Setup

1. Open the target Google Sheet.
2. Go to `Extensions > Apps Script`.
3. Paste the contents of:

```text
google_apps_script/Code.gs
```

4. Deploy it as a web app.
5. Set it to execute as you.
6. Copy the deployed web app URL.
7. Paste that URL into the app's `Apps Script Web App URL` field.
8. Paste the Google Sheet URL or ID into the app's `Google Sheet URL or ID` field.

Keep the Apps Script web app URL private because it can update any spreadsheet ID you pass to it.

### Does Sync Append?

No. Sync does not append rows to the bottom of an existing sheet.

It creates or replaces these named tabs in the target spreadsheet:

- `Settings 2026`
- `Raw Txns 2026`
- `Review <Month>`

If `Review June 2026` already exists, it is replaced with the newly generated version.

If you use a new label, such as `Review July 2026`, a new review tab is created.

This avoids stale rows, duplicate transaction imports, and broken monthly formulas.

The Apps Script sync writes values and formulas, freezes rows, and auto-resizes columns. For the full Excel formatting exactly as generated, use the downloaded `.xlsx` workbook or upload/import that workbook to Google Sheets manually.

## Google Sheet Housekeeping

Keep these active tabs:

- `Settings 2026`
- `Raw Txns 2026`
- current `Review <Month>` tabs

Old tabs from earlier experiments can be renamed with `Archive - ` or moved to the far right. Keep `Txns 25 Mar - 24 Apr 2026` if you like it as a visual reference.

## Troubleshooting

If the app does not start:

- Install Python 3.11 or newer
- On Windows, double-click `start_windows.bat` again
- If port `5057` is already in use, close the old app window/terminal and retry

If ABN upload fails:

- Confirm the file has `transactiondate`, `amount`, and `description` columns
- Use the original ABN export format if possible

If Wise upload fails:

- Export Wise transaction history as CSV or Excel
- Keep the default Wise column names

If formulas look wrong:

- Check `Raw Txns 2026` categories and `Include`
- Check the period label in `Review <Month>` cell `B3`
- Check that budgets exist in `Settings 2026`

If Google sync fails:

- Confirm the Apps Script web app URL is deployed and accessible
- Confirm the Google Sheet URL/ID is correct
- Re-authorize the Apps Script deployment if Google asks

## Development

Manual setup:

```sh
python -m venv .venv
.venv/Scripts/python -m pip install -r requirements.txt
.venv/Scripts/python app.py
```

On Mac/Linux, use `.venv/bin/python` instead of `.venv/Scripts/python`.
