# بيان المجموعات الدوائية — Pharmacy Drug-Group Report

A small static Arabic (RTL) form app for pharmacies to submit monthly drug-group totals by category (students / workforce / infants), plus a bulk file uploader.

This version keeps the original comic "ink outline" look: hard offset shadows, the zig-zag masthead, the Rx badge, the dotted paper and the per-category accent colours. It moves that look to a rounded, pill-shaped design and makes data entry faster.

## Files
- `index.html` — markup
- `css/style.css` — all styles (light/dark, per-category accent themes, list/grid views)
- `js/app.js` — all logic
- `images/logo.jpg` — brand mark and favicon

## What changed in this version (UX)
- **No more dropdowns.** Every choice can be seen and picked with one tap:
  - Year: a stepper (‹ 2025-2026 ›). Tap the value to go back to the current fiscal year, which is tagged "الحالية".
  - Region: a 3-way segmented control
  - Month: a 12-chip grid. A dot marks the suggested month (the previous month).
  - Pharmacy: chips for the chosen category and region. Pharmacies already sent this month have a green check and a dashed outline, and the label shows "X of Y sent".
  - The hidden native `<select>` elements are still the source of truth, so the saved draft, payload and storage keys are unchanged.
- **Reach any of the 18 groups straight away:**
  - Compact **list view** (the default): each group is a slim pill row, and all 18 fit in about one screen on desktop. You can switch to the **card grid view**; your choice is remembered.
  - **Search**: press `/` from anywhere, type a name (the search ignores Arabic letter variants like أ/ا and ة/ه) or a number, then press `Enter` to jump into the first match
  - **Filter**: All / Filled / Empty, each with a live count. Use "Empty" to see only what's left.
  - `↑ / ↓ / Enter / Shift+Enter` move between the visible fields. `Esc` goes back to the search box.
  - Click anywhere on a row to focus its field
  - The **progress ring** in the dock shows how many groups are filled. Tap it to jump to the next empty group.
- Number fields are wider and taller. Values show thousands separators when you leave a field (for example `987,654,321.125`) and go back to plain digits while you edit. The text shrinks to fit so large numbers are never cut off. The same applies to the grand total and the dock total.
- After a send, focus goes to the next pharmacy that hasn't been sent yet
- Kept from before: status chips that jump to missing items, auto-save, Reset with Undo, Ctrl/⌘+Enter to submit, pasting from Excel, Arabic-Indic digits, the "already sent" confirmation, bulk upload with sheet picker, and loading SheetJS only when needed

## Endpoint
- `GET https://firecloud.rubberylock7.workers.dev/submit/pharmacy-sheet?data=<json>` (`no-cors`, so the server never confirms receipt)

## Storage (localStorage)
`pharmacyAppData` (form draft, same format as before), `pharmacyAppTheme`, `pharmacyAppStatusHidden`, `pharmacyAppSent` (sent log per category/year/month), `pharmacyAppView` (list/grid)

## Not implemented / next steps
- Confirmation that the server received the data (needs a CORS-enabled endpoint that returns a response)
- Offline retry queue for failed sends
