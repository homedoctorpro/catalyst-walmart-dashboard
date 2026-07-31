# Catalyst Pet — Walmart Sales Dashboard

## Project Overview

Two-part analytics system for tracking Catalyst Pet product performance at Walmart:

1. **Weekly Sales Dashboard** (root) — Reads weekly Excel reports (currently 7 weeks: 202601–202607), computes KPIs, and generates a single self-contained HTML dashboard with interactive charts and maps.
2. **Store Distribution Map** (`Walmart Store List Map/`) — Separate pipeline that maps 3,735+ Walmart stores carrying Catalyst Pet SKUs, color-coded by SKU combination.

---

## Repository Structure

```
Walmart Sales Map/
├── extract_data.py             # Main ETL: Excel → JSON → dashboard.html
├── dashboard_template.html     # HTML/JS template (Plotly charts, Leaflet map, password gate)
├── dashboard.html              # Generated output (2.3 MB, do not edit manually)
├── stores_geo.json             # Geocoding cache: zip code → {lat, lon}
├── 202601-202607 *.xlsx        # Weekly input reports (7 files, do not commit to git)
└── Walmart Store List Map/
    ├── scripts/
    │   ├── 01_data_preparation.py
    │   ├── 02_geocoding.py
    │   ├── 03_generate_html.py
    │   └── utils.py            # SKU color mapping, JSON helpers
    ├── data/processed/         # Intermediate JSON files
    └── output/                 # Generated HTML maps
```

---

## Weekly Dashboard Workflow

### To regenerate `dashboard.html` after adding new Excel files:
```bash
python extract_data.py
```

### Pipeline (extract_data.py):
1. Read 2 sheets per Excel file (per `SHEET_MAP` dict — each week may have different sheet names)
   - **InStore Sales sheet**: rows 2–5 = 4 SKUs, row 6 = Total → POS $, units, instock %, U/S/W
   - **Sales by Store sheet**: per-store qty + on-hand inventory
2. Aggregate across weeks → per-store summaries, state-level rollups, consecutive OOS counts
3. Geocode zip codes via `pgeocode` (cache in `stores_geo.json`)
4. Embed JSON into `dashboard_template.html` (replacing `/*DATA_PLACEHOLDER*/`) → `dashboard.html`

### When the user says "run the script" or "process this week's data":

Do all of the following without being asked:

1. **Check what Excel files exist** (`ls *.xlsx`) and identify any new week codes not yet in `SHEET_MAP`.
2. **Get the sheet names** for each new file:
   `python -c "import openpyxl; wb = openpyxl.load_workbook('<filename>.xlsx', read_only=True); print(wb.sheetnames)"`
3. **Add entries** to `SHEET_MAP` in `extract_data.py` — match `instore`, `bystore`, `ecomm_l52`, `ecomm_lw` keys to actual sheet names (trailing spaces matter; copy exactly). Set any sheet that doesn't exist that week to `None`.
4. **Run** `python extract_data.py` — regenerates `dashboard.html`, sends the email (full distro on Mondays, dev-only otherwise). Week labels are auto-derived from the week code (`compute_week_label`), so no manual label edits are needed.
5. **Watch for `[WARN] Unmapped Catalyst ecomm row:` lines** in the output — that means Walmart changed a Catalyst product name in a way the token matcher couldn't handle. Update `parse_ecomm_product` in `extract_data.py` if it fires.
6. **If it's not Monday** and the user wants the full distro sent, run the manual send snippet from the session or use `python email_report.py` with `dev_only=False`.
7. **Commit and push** to deploy to GitHub Pages:
   `git add dashboard.html dashboard_template.html store_map.html store_map_template.html extract_data.py email_report.py stores_geo.json && git commit -m "Add week XXXXXX data" && git push origin master`

### When the calendar year rolls over (e.g. 202701 lands):
Add the new fiscal year's week-1 Friday to `FISCAL_YEAR_WEEK1_FRIDAY` in `extract_data.py` (one line, e.g. `"2027": date(2027, 2, 5)`). Until that's set, week labels for the new year fall back to the raw `YYYYWW` code.

### Adding a new week (manual reference):
1. Add the new Excel file to the root directory
2. Check the sheet names: `python -c "import openpyxl; wb = openpyxl.load_workbook('<filename>.xlsx', read_only=True); print(wb.sheetnames)"`
3. Add an entry to `SHEET_MAP` in `extract_data.py` with the correct sheet names
4. Run `python extract_data.py`
5. Commit and push to deploy: `git add dashboard.html dashboard_template.html store_map.html store_map_template.html extract_data.py email_report.py stores_geo.json && git commit -m "Add week XXXXXX data" && git push origin master`

### Hosting & Deployment
- **Hosted on**: GitHub Pages via `https://github.com/homedoctorpro/catalyst-walmart-dashboard`
- **Deployed branch**: `master` — GitHub Pages serves directly from this branch
- **To deploy**: commit the updated `dashboard.html` (+ `extract_data.py`, `stores_geo.json`) and push to `master`; GitHub Pages rebuilds automatically within ~1–2 minutes
- **Public store map (no password)**: `https://homedoctorpro.github.io/catalyst-walmart-dashboard/store_map.html` — standalone copy of the dashboard's 🏪 Store Map tab, generated as `store_map.html` from `store_map_template.html` by `write_store_map()` on every `extract_data.py` run. Safe to share externally; commit it with the weekly push.
- **Do not** push raw Excel files (`.xlsx`) to the repo

---

## Growth Report (`build_growth_report.py`)

Branded 2-slide PPTX + 2-sheet XLSX for stakeholder decks, output to `Growth Report/`:
1. **Subscriber growth by year** — Shopify DTC year-end active subscriptions (bar chart, YoY %).
2. **Walmart 15-lb Original (CATALYST15ORIG) linear weekly growth** — in-store units and
   U/S/W by FY26 Walmart week, each with a least-squares linear trend line.

### To refresh (when the user asks to update the growth report):
```bash
python build_growth_report.py
```
That's the whole refresh — no per-week config. It globs `2026?? Weekly Sales Report*.xlsx`
in the repo root and **auto-detects the in-store sheet** in each workbook (scans every sheet
for a `CATALYST15ORIG` row in column A, rows 1–6, requiring ≥34 columns; units = col H,
U/S/W = col AH). Watch for `[WARN] no 15O row in <week>` — that means a workbook's layout
changed and `find_15o_row()` needs adjusting.

Outputs land in `Growth Report/`: `Catalyst_Growth_Report.pptx`, `Catalyst_Growth_Report.xlsx`,
plus the chart PNGs (`rpt_subs.png`, `rpt_wm_units.png`, `rpt_wm_usw.png`). Commit the PPTX +
XLSX (the XLSX has a `.gitignore` exception; the PNGs are intermediates, don't commit).

### Notes / gotchas
- **Subscriber data is a static export**: `Growth Report/sub_growth_series.json` (monthly
  active-subscription counts from Recharge, exported from the CustomerIntelligence repo,
  data through Jun 2026). The subscriber slide only uses year-end (Dec) values for
  2022–2025, so it stays correct until end of 2026. To add 2026, re-export the series in
  CustomerIntelligence and add the year to the `years` list in `load_sub_growth()`.
- **When FY27 workbooks land** (`2027xx`, ~Feb 2027): update the glob pattern in
  `load_walmart_15o()` and the `FY26_W1_FRIDAY` constant / axis labels (or split by FY).
- History: this tool was built Jun 2026 in the CustomerIntelligence repo and moved here
  2026-07-31 (workbooks live here, so the refresh has no cross-repo dependency).

---

## Dashboard Features

- **Password gate**: `pellets123` (stored in `localStorage`)
- **5 tabs**: Overview, Sales Map, OOS Map, State Stats, State Charts
- **Charts** (Plotly.js): U/S/W, Adjusted U/S/W, Instock %, Units, Retail $, Wholesale $
- **Adjusted U/S/W**: `usw / (instock_pct / 100)` — normalizes by inventory availability
- **SKU chips**: filter which SKUs are shown across all charts simultaneously
- **Sales Map**: Leaflet.js dots, sqrt color scale, colored by consecutive OOS weeks
- **State heatmaps**: white (0%) → deep blue (100%) OOS fraction

---

## Data Structure (JSON embedded in dashboard)

```json
{
  "weeks": ["202601", ..., "202606"],
  "metrics": { "week": { "sku": { "pos_dollars", "pos_qty", "instock_pct", "usw", "wholesale_dollars" } } },
  "stores": { "store_num": { "state", "city", "zip", "lat", "lon" } },
  "weekly_stores": { "week": { "store_num": { "total_qty", "skus": { "sku": { "qty", "on_hand" } } } } },
  "state_sales": { "week": { "state": "total_qty" } },
  "state_oos": { "week": { "state": "oos_fraction" } },
  "consecutive_oos_by_week": { "week": { "sku": { "store_num": "count" } } }
}
```

---

## Store Distribution Map Workflow

Run scripts in order from `Walmart Store List Map/`:
```bash
python scripts/01_data_preparation.py   # Excel → stores_prepared.json
python scripts/02_geocoding.py          # Nominatim API (slow ~22 min, cached)
python scripts/03_generate_html.py      # → output/index.html + variants
```

Output: self-contained Leaflet.js maps at `output/index.html` (tabs: SKU map, simple map, data table).

---

## Key Implementation Notes

- **Sheet name variability**: Each week's Excel has inconsistent sheet names. `SHEET_MAP` in `extract_data.py` maps week → sheet names. Always check when adding new weeks.
- **Geocoding cache**: `stores_geo.json` persists zip → lat/lon. New zips are fetched and appended automatically.
- **Self-contained outputs**: Both `dashboard.html` and store map HTMLs embed all data inline — no external dependencies beyond CDN JS libs.
- **Lazy tab rendering**: Dashboard tabs only initialize charts on first click.
- **Consecutive OOS**: Counts trailing weeks where `on_hand == 0`; chain breaks if store is missing for a week.

---

## Technologies

| Layer | Tool |
|-------|------|
| Data processing | Python, pandas, pgeocode |
| Dashboard charts | Plotly.js 2.35.2 |
| Maps | Leaflet.js 1.9.4 + markercluster |
| Store map geocoding | Nominatim (OpenStreetMap, 1.1s rate limit) |
| Output format | Self-contained HTML (data embedded inline) |

---

## gstack Skills

**Web browsing:** Always use `/browse` for all web browsing tasks. Never use `mcp__claude-in-chrome__*` tools.

**Available skills:**

| Skill | Purpose |
|-------|---------|
| `/browse` | Browse the web using a headless Chromium browser |
| `/plan-ceo-review` | Generate a CEO-level review plan |
| `/plan-eng-review` | Generate an engineering review plan |
| `/review` | Code/PR review |
| `/ship` | Ship a feature end-to-end |
| `/qa` | QA a feature or change |
| `/qa-only` | Run QA checks only (no planning) |
| `/setup-browser-cookies` | Configure browser session cookies |
| `/retro` | Run a retrospective |
| `/document-release` | Document a release |
