# SDAIA Training Programs 2023 — Data Insight Dashboard

A standalone, browser-based analytics dashboard for the Institute of Information Technology’s 2023 training-program records from Saudi Open Data / SDAIA.

## Project goal

The dashboard turns local 2023 training-program records into decision-ready comparisons of training terms, reported participation, available course volume, program distribution, and duration patterns. It reports workbook measures only; it does not establish causes, attendance, completion, outcomes, satisfaction, or unique-person counts.

## Highlights

- Five accessible dashboard views: Overview, Term Comparison, Program Participation, Duration, and Participation.
- English-first program labels with the exact original Arabic source title retained for traceability.
- Search program records in English or Arabic, with term and duration filters.
- Interactive charts with keyboard-accessible tooltips and textual data fallbacks.
- Dark mode by default, with a saved light/dark preference.
- Responsive design and reduced-motion support.
- No framework, server, or external dependency is required.

## Source and method

- **Primary source:** `xlsx file.xlsx`, a local Excel workbook containing detailed Institute of Information Technology 2023 records from four training terms.
- **Parity export:** `csv file.csv` is retained for reference and is not added to the dashboard calculations.
- The dashboard contains 64 normalized detail rows, totalling 14,575 reported trainee counts and 468 available course counts. One course-count source cell is blank and remains missing.

## Run locally

Open `index.html` directly in a browser, or serve this folder locally:

```bash
python -m http.server 4173 --bind 127.0.0.1
```

Then open `http://127.0.0.1:4173/`.

## Files

- `index.html` — semantic dashboard shell and source attribution
- `style.css` — responsive light/dark visual system
- `script.js` — embedded normalized data, calculations, charts, filtering, and interactions
