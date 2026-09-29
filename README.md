# بيان المجموعات الدوائية — Pharmacy Drug-Group Report

A small static Arabic (RTL) form app for pharmacies to submit monthly drug-group totals by category (students / workforce / infants), plus a bulk file uploader. It keeps the original comic "ink outline" look and gives it a calmer, more refined finish.

## Files
- `index.html` — markup
- `css/style.css` — all styles (light/dark, per-category accent themes)
- `js/app.js` — all logic

## Features
- 4-step form: category → pharmacy details → 18 drug groups → optional totals, with a big animated grand total
- Clickable status chips jump to the next missing field
- Sticky action dock showing the live total, auto-save indicator, **Reset** (with Undo) and **Submit** (Ctrl/⌘ + Enter)
- After a successful send, category, year, month and region are kept so you can pick the next pharmacy right away
- Pharmacies already sent for the same category and month on this device are marked "✓ أُرسلت", and you're asked to confirm before sending again
- Each group shows its share of the total with a small meter; each has a quick-clear button with Undo
- Enter / Shift+Enter moves between number fields; pasting a column or row from Excel fills the following groups
- Arabic-Indic digits and Arabic decimal separators are accepted
- Light / Dark / System theme; accent colour follows the chosen category, applied before first paint
- Bulk upload: drag and drop anywhere on the page, choose Excel sheets (select all / none), file size and type badges, a Stop button during sending, and a JSON backup of the rows that were sent
- The Excel library (SheetJS) loads only when the upload panel is opened

## Endpoint
- `GET https://firecloud.rubberylock7.workers.dev/submit/pharmacy-sheet?data=<json>` (`no-cors`, so the server never confirms receipt)

## Storage (localStorage)
`pharmacyAppData` (form draft), `pharmacyAppTheme`, `pharmacyAppStatusHidden`, `pharmacyAppSent` (sent log per category/year/month)

## Not implemented / next steps
- Confirmation that the server received the data (needs a CORS-enabled endpoint that returns a response)
- `favicon.ico` is referenced but not included in the project
- Optional: offline retry queue for failed sends
