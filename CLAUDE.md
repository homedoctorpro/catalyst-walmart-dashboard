# Catalyst Pet — Walmart Sales Dashboard

## Project Overview

Two-part analytics system for tracking Catalyst Pet product performance at Walmart:

1. **Weekly Sales Dashboard** (root) — Reads 6 weekly Excel reports, computes KPIs, and generates a single self-contained HTML dashboard with interactive charts and maps.
2. **Store Distribution Map** (`Walmart Store List Map/`) — Separate pipeline that maps 3,735+ Walmart stores carrying Catalyst Pet SKUs, color-coded by SKU combination.

---

## Repository Structure

```
Walmart Sales Map/
├── extract_data.py             # Main ETL: Excel → JSON → dashboard.html
├── dashboard_template.html     # HTML/JS template (Plotly charts, Leaflet map, password gate)
├── dashboard.html              # Generated output (2.3 MB, do not edit manually)
├── stores_geo.json             # Geocoding cache: zip code → {lat, lon}
├── 202601-202606 *.xlsx        # Weekly input reports (6 files)
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

### Adding a new week:
- Add the new Excel file to the root directory
- Add an entry to `SHEET_MAP` in `extract_data.py` with the correct sheet names
- Run `python extract_data.py`

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
