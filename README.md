# Google Apps Script samples

Small, working Google Sheets automations. Each folder is one script you can paste into Extensions -> Apps Script of a copy of your sheet.

| Folder | What it does |
|---|---|
| [page-change-alert](page-change-alert) | Watches public web pages and sends a Telegram or email alert when text appears or disappears, a number crosses a threshold or the page changes. One alert per change, 3 failed checks -> one error alert, every check logged to a sheet. Setup guide in `SETUP.md`. |
| [api-to-sheets](api-to-sheets) | Pulls a paged JSON API (Stripe charges as the example) into a sheet every hour. Updates existing rows by ID, adds new ones, never duplicates. The API key lives in Script Properties, not in the sheet. |
| [form-to-pdf-email](form-to-pdf-email) | New Google Form answer -> filled Google Docs template -> PDF in Drive -> emailed to the person, status written back to the sheet. Safe to re-run. |
| [merge-files](merge-files) | Merges the same tab from several spreadsheets into one master sheet, matching columns by header name, adds a Source column, drops duplicates and rebuilds a monthly summary. |

## Principles used in every script

- Batch reads and writes (`getValues` / `setValues`), never cell by cell.
- Secrets in Script Properties or a private Settings tab, never in code.
- `setup()` removes old triggers before creating new ones, so running it twice does not double the trigger.
- Errors are reported once, not on every run.

## License

MIT
