# بيان المجموعات الدوائية — Pharmacy Drug-Group Report

A small static Arabic (RTL) form app. Pharmacies use it to submit monthly drug-group totals by category (students / workforce / infants). It also includes a bulk file uploader.

Source: https://github.com/datakid/finalapp

## Design ("pharmacist's notebook")
Warm paper with fine grain, thin ink lines, hand-drawn loops, the Aref Ruqaa / IBM Plex Sans Arabic / Fraunces type system, and muted category accents (ink blue / terracotta / sage). Light and dark themes are both tuned. Animations are turned off when "reduce motion" is on.

### Maturity pass (this revision)
- **New mark and favicon** (`favicon.svg`): an abstract open ink loop with a single leaf inside it and a warm dot. There are no letters in it. In the header it draws itself in and the leaf takes the colour of the active category. The favicon follows the system light/dark setting.
- **Year rail** replaces the old stepper. It shows three fiscal years side by side (`2025–2026`), and the selected one gets the same pen loop as the other choices. On a new sheet the current fiscal year is already selected.
  - Click a year, use `← →`, or type the last two digits (`26` selects 2025–2026)
  - The side arrows slide the window to older or newer years
  - A "return to current year" link appears when you are on any other year
  - Stored values are unchanged (the `year` field holds the end year, e.g. `2026`)
- **Quick-entry line**: a large ruled line that stays at the top of the groups section while you scroll. It always shows the current target group.

  | Type | Effect |
  |---|---|
  | `4500` ↵ | sets the target group and moves to the next one |
  | `1200+350-50` ↵ | adds/subtracts the numbers and stores the result |
  | `+300` / `-50` ↵ | adds to / subtracts from the current value |
  | `قلب 4500` ↵ | sets a group by part of its name |
  | `5=4500` ↵ | sets a group by its number |
  | `قلب` ↵ (or Tab) | moves to that group; typing a name also filters the ledger |
  | ↵ on an empty line | skips to the next group (Shift+↵ goes back) |
  | `↑ ↓` | moves the target |
  | `Ctrl Z` | undoes the last entry |

  A live preview shows what Enter will do (`= 1,500`, or the old value it will replace) before you commit. Arabic-Indic digits and `٫` are accepted. Commas are read sensibly: `1,234,567` is thousands, `12,5` is a decimal. If you start typing digits anywhere on the page they go into the line. On touch screens a key switches the keyboard between numbers and letters.
- **Big numbers are readable**: numbers get thousands separators while you type, in both the quick line and the ledger cells, and the cursor stays in place. Group inputs are larger (1.5rem), and pasted Excel columns go through the same number parser.
- **Pickers are keyboard-fast**: type-ahead on month/region/pharmacy (letters match the name, digits match the position, e.g. `3` = March). `1 2 3` picks a category. Enter confirms and jumps to the next unfilled field, then to the quick-entry line.
- **Readability**: base text is now 16.5px, and labels, hints, chips and notes are larger. The muted greys have more contrast.

## Files
- `index.html`: markup, inline SVG mark, illustrations and filters
- `css/style.css`: all styles
- `js/app.js`: all logic
- `favicon.svg`: icon (replaces `favicon.ico` / `images/logo.jpg`, which are no longer referenced)

## Behaviour kept
All IDs used by the payload, auto-save, Reset with Undo, the "already sent" log, the `/` shortcut, `Ctrl+Enter` submit, Excel paste, the ledger/grid views, filters and bulk upload work as before.

## Endpoint
- `GET https://firecloud.rubberylock7.workers.dev/submit/pharmacy-sheet?data=<json>` (`no-cors`, so the server never confirms receipt)

## Storage (localStorage, keys unchanged, existing data loads as before)
`pharmacyAppData`, `pharmacyAppTheme`, `pharmacyAppStatusHidden`, `pharmacyAppSent`, `pharmacyAppView`

## Next steps
- Confirmation that the server received the data (needs a CORS-enabled endpoint)
- Offline retry queue for failed sends
- Optional PNG apple-touch-icon (iOS ignores SVG touch icons)
