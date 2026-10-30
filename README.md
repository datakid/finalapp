# بيان المجموعات الدوائية — Pharmacy Drug-Group Report

A small static Arabic (RTL) form app. Pharmacies use it to submit monthly drug-group totals by category (students / workforce / infants). It also includes a bulk file uploader.

Source: https://github.com/datakid/finalapp

## Design ("pharmacist's notebook")
The visuals were redone from scratch. The old comic look (hard offset shadows, zig-zag masthead) is gone. The new direction is a quiet hand-drawn notebook: warm paper with fine grain, thin ink lines and plenty of white space.

- **Type**: Aref Ruqaa (calligraphic headings), IBM Plex Sans Arabic (body text), Fraunces (numbers, set in italic for indexes)
- **Hand-drawn feel**: an SVG `feTurbulence` filter (`#sketch`, `#sketch-soft`) adds a slight wobble to card outlines, uneven border-radius values make shapes look drawn by hand, and selected options get a loop that looks drawn around them with a pen
- **Masthead**: a line illustration (mortar and pestle, capsule, herb sprig) and an underline flourish that draw themselves in when the page loads
- **Groups**: the default "دفتر" (ledger) view lays the 18 groups out in two columns. Each row reads name, dotted leader, then value, like a ledger, with a thin meter line for its share. The card grid view is still available.
- **Steps**: a pen-drawn tick appears beside each section heading once that section is complete
- **Category accents** are kept but muted: ink blue (students), terracotta (workforce) and sage (infants). Light and dark themes are both tuned.
- Animations are turned off when the system "reduce motion" setting is on

## Files
- `index.html` — markup plus inline SVG illustrations and filters
- `css/style.css` — all styles (rewritten)
- `js/app.js` — all logic (unchanged)
- `images/logo.jpg` — brand mark and favicon

## Behaviour (unchanged)
All IDs, class hooks, keyboard shortcuts (`/`, `↑ ↓`, `Enter`, `Ctrl+Enter`, `Esc`), Excel paste, auto-save, Reset with Undo, the "already sent" log and bulk upload work exactly as before.

## Endpoint
- `GET https://firecloud.rubberylock7.workers.dev/submit/pharmacy-sheet?data=<json>` (`no-cors`, so the server never confirms receipt)

## Storage (localStorage, keys unchanged)
`pharmacyAppData`, `pharmacyAppTheme`, `pharmacyAppStatusHidden`, `pharmacyAppSent`, `pharmacyAppView`

## Next steps
- Confirmation that the server received the data (needs a CORS-enabled endpoint)
- Offline retry queue for failed sends
