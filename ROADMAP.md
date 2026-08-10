# Roadmap

Ideas to improve the tool after the first working version.

## Web-Native Review

- Show the generated review directly inside the web app.
- Display the budget overview, income/credits, and category-grouped transactions as tables in the browser.
- Keep Excel export as an option, but do not require Excel or Google Sheets for normal monthly review.
- Let users adjust categories and include/exclude flags in the browser before exporting or syncing.

## Source-Agnostic Imports

- Accept any reasonable bank statement upload, not only ABN AMRO and Wise exports.
- Support Excel, CSV, and PDF statement uploads.
- Detect columns, dates, amounts, descriptions, and transaction direction automatically where possible.
- Add a review step for unknown formats so the user can confirm which columns mean date, amount, merchant, and description.
- Keep source-specific parsers for known banks, but fall back to a generic importer.

## Offline Analysis

- Explore optional offline analysis with a local Ollama model.
- Use the local model to suggest categories, detect recurring transactions, summarize month-over-month changes, and flag unusual spending.
- Keep deterministic rules and formulas as the source of truth; use AI suggestions as editable review helpers.
- Ensure financial data can stay local on the user's machine when offline mode is enabled.

## Privacy And Safety

- Avoid requiring users to upload bank data to third-party services for basic use.
- Keep generated files, uploads, and local configuration out of Git.
- Add clear warnings before any workflow that would store bank exports in GitHub, Google Drive, or another cloud service.
