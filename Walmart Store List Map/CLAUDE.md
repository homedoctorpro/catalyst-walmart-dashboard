# Catalyst Pet Store Map

## Project Overview

This project generates an interactive web map showing the distribution of **3,735 Walmart stores** carrying Catalyst Pet products across the United States. Stores are color-coded by which SKU combinations they carry. The output is a self-contained HTML file that can be hosted on GitHub Pages, Netlify, or opened locally.

## Project Structure

```
Walmart Store List Map/
├── Catalyst Store List.xlsx      # Source data - Walmart store/SKU list
├── CLAUDE.md                     # This file
├── scripts/                      # Python pipeline (run in order)
│   ├── utils.py                  # Shared utilities, SKU color mapping, JSON I/O
│   ├── 01_data_preparation.py    # Step 1: Excel → JSON (groups stores by SKU combos)
│   ├── 02_geocoding.py           # Step 2: City/State → lat/lon via Nominatim API
│   └── 03_generate_html.py       # Step 3: Geocoded JSON → self-contained HTML
├── data/
│   └── processed/
│       ├── stores_prepared.json   # Output of Step 1
│       ├── stores_geocoded.json   # Output of Step 2
│       └── geocoding_cache.json   # Nominatim API response cache
└── output/
    ├── index.html                 # Main output - tabbed HTML with map + data table
    ├── catalyst-store-map.html    # SKU-colored map (standalone)
    ├── catalyst-store-map-simple.html  # Single-color map (standalone)
    ├── data-table.html            # State summary table (standalone)
    ├── interactive-map.html       # Alternative map variant
    ├── README.md                  # End-user documentation
    ├── HOSTING-GUIDE.txt          # Instructions for web hosting
    └── GITHUB-PAGES-STEPS.txt     # Step-by-step GitHub Pages guide
```

## Pipeline

The scripts run sequentially — each step depends on the previous:

1. **`01_data_preparation.py`** — Reads `Catalyst Store List.xlsx`, extracts store number/name/city/state/item columns, groups by store, and builds SKU combination strings. Outputs `stores_prepared.json`.

2. **`02_geocoding.py`** — Takes prepared stores, deduplicates by city/state, geocodes unique locations via the Nominatim (OpenStreetMap) API with 1.1s rate limiting. Uses a file-based cache (`geocoding_cache.json`) to avoid repeat API calls. Outputs `stores_geocoded.json`.

3. **`03_generate_html.py`** — Reads geocoded stores, generates GeoJSON, creates state-level summaries, and produces self-contained HTML files using Leaflet.js with embedded data. The main output is `index.html` with tabs for SKU Map, Simple Map, and Data Table.

## Key Data

- **4 SKU products**: CATALYST15ORIG, CATALYST15UNSCEN, CATALYST34LBORIGINAL, CATALYSTPET34LBUNSCE
- **11 unique SKU combinations** across stores (single, pairs, triples, all-four)
- Each combination has a designated color defined in `utils.py:SKU_COLORS`
- SKU abbreviations used in display: 15O, 15U, 34O, 34U

## Dependencies

- Python 3
- `pandas` — Excel reading and data manipulation
- `requests` — HTTP calls to Nominatim API
- `tqdm` — Progress bars for geocoding
- `openpyxl` — Excel file support (pandas backend)

## Frontend Stack

- **Leaflet.js 1.9.4** — Map rendering with OpenStreetMap tiles
- **Leaflet.markercluster** — Marker clustering at low zoom levels
- All HTML output files are self-contained (no local asset dependencies)

## Important Notes

- The geocoding step respects Nominatim's rate limit (1.1s between requests). A full geocoding run of ~1,200 unique locations takes ~22 minutes.
- The geocoding cache persists between runs. Delete `geocoding_cache.json` to force re-geocoding.
- The Excel source file columns expected: `store_number`, `store_name`, `city_name`, `state_or_province_name`, `item_name`.
- Output HTML files embed all store data inline as JavaScript — no external data files needed.
