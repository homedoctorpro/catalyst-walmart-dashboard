# Catalyst Pet Store Distribution Map

Interactive map visualization of 3,735 Walmart stores carrying Catalyst Pet products across the United States.

## Quick Start

1. Open `index.html` in any web browser to view the interactive map
2. Open `data-table.html` to view store counts by state

## Features

### Interactive Map (index.html)
- **Color-coded markers** by SKU combination (11 unique combinations)
- **Zoom and pan** to explore store locations
- **Click markers** for detailed store information
- **Filter by combination** - click legend items to filter
- **Search** by store number, city, or state
- **Marker clustering** for better performance at low zoom levels

### Data Table (data-table.html)
- **Sortable columns** - click headers to sort
- **Filter by state** - search box for quick filtering
- **Store counts** by SKU across all states
- **Total row** showing aggregated statistics

## Data Summary

- **Total Stores**: 3,734
- **States Covered**: 50
- **SKU Products**: 4
  - CATALYST15ORIG
  - CATALYST15UNSCEN
  - CATALYST34LBORIGINAL
  - CATALYSTPET34LBUNSCE
- **Unique Combinations**: 11
- **Generated**: 2026-02-10 16:37:32

## Technical Details

- **Mapping**: Leaflet.js with OpenStreetMap tiles
- **Geocoding**: Nominatim API
- **Format**: Self-contained HTML files (works offline after initial load)
- **Browser Support**: Chrome, Firefox, Safari, Edge (latest versions)

## Sharing

This map is fully self-contained. Simply:
1. Share the entire folder (including all files)
2. Recipient opens `index.html` in their browser
3. Works on Windows, Mac, Linux

No server or installation required!

## Files

- `index.html` - Interactive map page
- `data-table.html` - State summary table
- `README.md` - This file

---

Generated with Catalyst Pet Store Mapper
