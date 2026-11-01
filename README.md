# بيان المجموعات الدوائية — Pharmacy Drug-Group Report

A small static Arabic (RTL) form app. Pharmacies use it to submit monthly drug-group totals by category (students / workforce / infants). It also includes a bulk file uploader.

Source: https://github.com/datakid/finalapp

## Design ("pharmacist's notebook")
Warm paper with fine grain, thin ink lines, hand-drawn loops, the Aref Ruqaa / IBM Plex Sans Arabic / Fraunces type system, and muted category accents (ink blue / terracotta / sage). Light and dark themes are both tuned. Animations are turned off when "reduce motion" is on.

## Two ways to fill the form
The form and the command palette write through the same actions, so history, auto-save, undo and the "already sent" marks behave identically from either.

### The form
Pick the category, type the fiscal year, choose month, region and pharmacy, then fill the 18 groups. `Enter` and the arrow keys move between fields; `Esc` in a group input leaves it.

### The command palette (the fast way)
Open it with `Ctrl K` / `⌘ K`, `/`, the search button (top bar, and the dock on phones), or by typing a digit anywhere on the page. It is the only search and command surface.

| Type | Result |
|---|---|
| `4500` ↵ | sets the current group; the next empty group becomes current and the palette stays open |
| `قلب 4500` ↵ | sets a group by part of its name |
| `قلب +300`, `قلب 1200+350` ↵ | adds to the value, or sums |
| `5=4500`, `#5 4500` ↵ | sets a group by its number |
| `قلب 4500؛ كلي 300` ↵ | sets several groups at once (`؛`, `;` or a new line; never a comma). One `Ctrl Z` reverts the whole line. An invalid part is named and nothing is applied |
| `2024`, `2024-2025`, `24/25`, `سنة 2024` ↵ | sets the fiscal year (see below) |
| `مارس`, `شهر 3` ↵ | sets the month |
| `دمنهور` ↵ | finds pharmacies across all categories and regions; the current ones rank first, and choosing one also switches category and region |
| `تذاكر 1200`, `خارجي 300` ↵ | sets the optional totals |
| `فارغة`, `داكن`, `بطاقات`, `إرسال`, `رفع` ↵ | commands |
| empty | next missing field, the current group, every group with its value, then commands |

Filling all 18 groups takes only digits and `Enter`. `Shift+↵` on an empty line moves back one group, `Ctrl Z` undoes the last entry (a whole batch counts as one), `Tab` completes a name, and `↑ ↓` move the highlight. The highlighted row is always what `Enter` does. Arabic-Indic digits and `٫` are accepted, and commas are read sensibly (`1,234,567` is thousands, `12,5` is a decimal).

A bare four-digit number within ten years of the current fiscal year ranks the year first and the group second; otherwise the group comes first. A separator or the word سنة always means a year. On phones the palette is a bottom sheet that stays above the keyboard, with a `123 / أ` key to switch keyboards.

## Fiscal year
Type the start year and the field completes the season: `2024` becomes `2024 – 2025`.

- Any start year from 1900 to 2198 is accepted; the stored `year` is still the end year, exactly as before
- Two digits expand with a rolling window (`24` is 2024, `98` is 1998); `24/25`, `2024-2025`, `20242025` and Arabic-Indic digits work
- A wrong end year (`2024-2027`) or an out-of-range year shows a message and is never saved
- Years more than ten away from the current fiscal year ask for confirmation (a second `Enter` or the "yes" link); leaving the field without confirming restores the previous year
- `↑ ↓` step by one year, `Shift+↑ ↓` or `PageUp/PageDown` by ten, `Home` returns to the current year, `Esc` reverts the edit
- The current and previous fiscal years are always one tap away

## Groups
Numbers accept up to `1,000,000,000.000` and get thousands separators while you type, with the cursor staying in place. Inputs keep one comfortable size and only step down when a value would not fit. Values of a million and over show a muted `≈ 1.25 مليون` hint. Pasted Excel columns go through the same number parser. The ledger (دفتر) is one column below 1100px; the card view (بطاقات) keeps equal-height cards.

## Keyboard map
| Key | Where | Action |
|---|---|---|
| `Ctrl K` / `⌘ K`, `/`, a digit | page | open the palette |
| `Ctrl Z` | outside text fields | undo the last entry |
| `Ctrl Enter` | anywhere | submit |
| `Enter`, `↓`, `↑` | fields | next / previous field |
| `Esc` | group input | leave the field |
| `Esc` | palette | clear the line, then close |

## Files
- `index.html`: markup, inline SVG mark, illustrations and filters
- `css/style.css`: all styles, built on spacing, tap-target and motion tokens
- `js/core.js`: pure helpers shared by the form and the palette (number parsing, entry grammar, fiscal-year rules, scoring)
- `js/app.js`: state, the actions every writer goes through, rendering and the palette
- `tests/core.test.html`: open it in a browser; it runs the core helpers against a table of cases and shows a pass/fail list
- `favicon.svg`: icon

## Behaviour kept
The submit payload (same keys, same order, `year` is the end year), auto-save, Reset with Undo, the "already sent" log, `Ctrl+Enter` submit, Excel paste, the ledger/grid views, filters, theme and bulk upload work as before.

## Endpoint
- `GET https://firecloud.rubberylock7.workers.dev/submit/pharmacy-sheet?data=<json>` (`no-cors`, so the server never confirms receipt)

## Storage (localStorage, keys unchanged, existing data loads as before)
`pharmacyAppData`, `pharmacyAppTheme`, `pharmacyAppStatusHidden`, `pharmacyAppSent`, `pharmacyAppView`

## Conventions
No code comments in JS, CSS or HTML. Commits use `type: one-line summary`.

## Next steps
- Confirmation that the server received the data (needs a CORS-enabled endpoint)
- Offline retry queue for failed sends
- Optional PNG apple-touch-icon (iOS ignores SVG touch icons)
