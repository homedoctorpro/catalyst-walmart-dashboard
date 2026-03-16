"""
Step 3: HTML Generation
Creates self-contained HTML files with embedded data for interactive map and data table
"""
import sys
import json
from pathlib import Path
from collections import defaultdict
from datetime import datetime

# Add scripts directory to path for imports
sys.path.append(str(Path(__file__).parent))

from utils import load_json, save_json, get_project_root, SKU_COLORS


def format_sku_display(sku_combination):
    """
    Convert SKU combination to readable abbreviation format

    Args:
        sku_combination: Comma-separated SKU string

    Returns:
        Formatted display string (e.g., "15O / 34O")
    """
    sku_abbrev = {
        'CATALYST15ORIG': '15O',
        'CATALYST15UNSCEN': '15U',
        'CATALYST34LBORIGINAL': '34O',
        'CATALYSTPET34LBUNSCE': '34U'
    }

    skus = sku_combination.split(', ')
    abbreviated = [sku_abbrev.get(sku, sku) for sku in skus]
    return ' / '.join(abbreviated)


def create_geojson(stores):
    """
    Convert stores to GeoJSON format

    Args:
        stores: List of store dictionaries

    Returns:
        GeoJSON FeatureCollection
    """
    features = []

    for store in stores:
        # Only include stores with valid coordinates
        if store.get('latitude') and store.get('longitude'):
            feature = {
                "type": "Feature",
                "geometry": {
                    "type": "Point",
                    "coordinates": [store['longitude'], store['latitude']]
                },
                "properties": {
                    "store_number": store['store_number'],
                    "store_name": store['store_name'],
                    "street_address": store.get('street_address', ''),
                    "city": store['city'],
                    "state": store['state'],
                    "sku_combination": store['sku_combination'],
                    "skus": store['skus'],
                    "sku_count": store['sku_count'],
                    "color": store['color']
                }
            }
            features.append(feature)

    geojson = {
        "type": "FeatureCollection",
        "features": features,
        "metadata": {
            "total_stores": len(stores),
            "geocoded_stores": len(features),
            "failed_geocoding": len(stores) - len(features),
            "generated_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
    }

    return geojson


def create_state_summary(stores):
    """
    Create state-level summary statistics

    Args:
        stores: List of store dictionaries

    Returns:
        State summary dictionary
    """
    # Initialize state data
    state_data = defaultdict(lambda: {
        'total_stores': 0,
        'sku_counts': {
            'CATALYST15ORIG': 0,
            'CATALYST15UNSCEN': 0,
            'CATALYST34LBORIGINAL': 0,
            'CATALYSTPET34LBUNSCE': 0
        }
    })

    # Aggregate by state
    for store in stores:
        state = store['state']
        state_data[state]['total_stores'] += 1

        # Count individual SKUs
        for sku in store['skus']:
            if sku in state_data[state]['sku_counts']:
                state_data[state]['sku_counts'][sku] += 1

    # Convert to list format and sort by total stores descending
    states = []
    sorted_states = sorted(state_data.items(), key=lambda x: x[1]['total_stores'], reverse=True)
    for state_name, data in sorted_states:
        states.append({
            'state': state_name,
            'total_stores': data['total_stores'],
            'sku_counts': data['sku_counts']
        })

    # Calculate totals
    totals = {
        'total_stores': sum(s['total_stores'] for s in states),
        'CATALYST15ORIG': sum(s['sku_counts']['CATALYST15ORIG'] for s in states),
        'CATALYST15UNSCEN': sum(s['sku_counts']['CATALYST15UNSCEN'] for s in states),
        'CATALYST34LBORIGINAL': sum(s['sku_counts']['CATALYST34LBORIGINAL'] for s in states),
        'CATALYSTPET34LBUNSCE': sum(s['sku_counts']['CATALYSTPET34LBUNSCE'] for s in states)
    }

    return {
        'states': states,
        'totals': totals,
        'generated_date': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }


def generate_map_html(geojson, state_summary):
    """Generate the interactive map HTML"""

    html = '''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Catalyst Pet Store Distribution Map</title>

    <!-- Leaflet CSS -->
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css" />
    <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css" />

    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }

        #header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px 30px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }

        #header h1 {
            font-size: 28px;
            margin-bottom: 8px;
        }

        #header p {
            font-size: 14px;
            opacity: 0.95;
        }

        .nav-links {
            margin-top: 12px;
        }

        .nav-links a {
            color: white;
            text-decoration: none;
            margin-right: 15px;
            padding: 8px 16px;
            border: 1px solid rgba(255,255,255,0.5);
            border-radius: 4px;
            transition: all 0.3s;
            display: inline-block;
            font-size: 14px;
        }

        .nav-links a:hover {
            background: white;
            color: #667eea;
        }

        .nav-links a.active {
            background: rgba(255,255,255,0.2);
            border-color: white;
        }

        #map {
            height: calc(100vh - 140px);
            width: 100%;
        }

        .legend {
            background: white;
            padding: 12px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
            max-height: 500px;
            overflow-y: auto;
            min-width: 200px;
        }

        .legend h4 {
            margin: 0 0 10px 0;
            font-size: 14px;
            color: #333;
            border-bottom: 2px solid #667eea;
            padding-bottom: 6px;
        }

        .legend-item {
            display: flex;
            align-items: center;
            margin: 6px 0;
            cursor: pointer;
            padding: 6px;
            border-radius: 4px;
            transition: background 0.2s;
            font-size: 12px;
        }

        .legend-item:hover {
            background: #f5f5f5;
        }

        .legend-item.filtered {
            background: #e3f2fd;
            border: 1px solid #2196F3;
        }

        .legend-color {
            width: 18px;
            height: 18px;
            border-radius: 50%;
            margin-right: 8px;
            border: 2px solid #333;
            flex-shrink: 0;
        }

        .legend-text {
            font-size: 11px;
            line-height: 1.3;
        }

        .search-box {
            position: absolute;
            top: 10px;
            right: 10px;
            z-index: 1000;
            background: white;
            padding: 10px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
        }

        .search-box input {
            width: 250px;
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 13px;
        }

        .search-box input:focus {
            outline: none;
            border-color: #667eea;
        }

        .leaflet-popup-content {
            font-size: 13px;
            line-height: 1.5;
        }

        .leaflet-popup-content strong {
            color: #667eea;
        }

        @media (max-width: 768px) {
            #map { height: calc(100vh - 160px); }
            .search-box input { width: 180px; }
            .legend { max-height: 300px; }
        }
    </style>
</head>
<body>
    <div id="header">
        <h1>Catalyst Pet - Walmart Store Distribution</h1>
        <p>''' + f"{geojson['metadata']['geocoded_stores']:,}" + ''' stores across 50 states | 4 SKUs | 11 unique combinations</p>
        <div class="nav-links">
            <a href="index.html" class="active">Map View</a>
            <a href="data-table.html">Data Table</a>
        </div>
    </div>

    <div id="map"></div>

    <!-- Leaflet JS -->
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <script src="https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js"></script>

    <!-- Embedded Store Data -->
    <script>
    const STORE_DATA = ''' + json.dumps(geojson, separators=(',', ':')) + ''';

    const SKU_COLORS = ''' + json.dumps(SKU_COLORS, separators=(',', ':')) + ''';
    </script>

    <script>
        // Format SKU display with abbreviations
        function formatSkuDisplay(skuCombination) {
            const skuAbbrev = {
                'CATALYST15ORIG': '15O',
                'CATALYST15UNSCEN': '15U',
                'CATALYST34LBORIGINAL': '34O',
                'CATALYSTPET34LBUNSCE': '34U'
            };

            const skus = skuCombination.split(', ');
            const abbreviated = skus.map(sku => skuAbbrev[sku] || sku);
            return abbreviated.join(' / ');
        }

        // Initialize map
        const map = L.map('map').setView([39.8283, -98.5795], 5); // Center of USA

        // Add OpenStreetMap tiles
        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
            maxZoom: 18
        }).addTo(map);

        // Create marker cluster group
        const markers = L.markerClusterGroup({
            maxClusterRadius: 50,
            spiderfyOnMaxZoom: true,
            showCoverageOnHover: false,
            disableClusteringAtZoom: 10
        });

        // Store markers by combination for filtering
        const markerLayers = {};

        // Add markers
        STORE_DATA.features.forEach(feature => {
            const props = feature.properties;
            const coords = feature.geometry.coordinates;

            if (coords[1] && coords[0]) {
                const marker = L.circleMarker([coords[1], coords[0]], {
                    radius: 6,
                    fillColor: props.color,
                    color: '#000',
                    weight: 1,
                    opacity: 1,
                    fillOpacity: 0.8
                });

                marker.bindPopup(`
                    <strong>${props.store_name}</strong><br>
                    Store #${props.store_number}<br>
                    ${props.street_address}<br>
                    ${props.city}, ${props.state}<br>
                    <br>
                    <strong>SKUs (${props.sku_count}):</strong><br>
                    ${props.skus.map(sku => '• ' + sku).join('<br>')}
                `);

                if (!markerLayers[props.sku_combination]) {
                    markerLayers[props.sku_combination] = [];
                }
                markerLayers[props.sku_combination].push(marker);
                markers.addLayer(marker);
            }
        });

        map.addLayer(markers);

        // Add legend with filtering
        const legend = L.control({position: 'bottomleft'});
        legend.onAdd = function(map) {
            const div = L.DomUtil.create('div', 'legend');
            div.innerHTML = '<h4>SKU Combinations</h4>';

            // Sort combinations by frequency
            const comboCounts = {};
            STORE_DATA.features.forEach(f => {
                const combo = f.properties.sku_combination;
                comboCounts[combo] = (comboCounts[combo] || 0) + 1;
            });

            const sortedCombos = Object.entries(comboCounts).sort((a, b) => b[1] - a[1]);

            sortedCombos.forEach(([combo, count]) => {
                const color = SKU_COLORS[combo];
                const displayName = formatSkuDisplay(combo);

                div.innerHTML += `
                    <div class="legend-item" onclick="filterMarkers('${combo}')" id="legend-${combo.replace(/[, ]/g, '_')}">
                        <div class="legend-color" style="background-color: ${color}"></div>
                        <div class="legend-text"><strong>${count}</strong> - ${displayName}</div>
                    </div>
                `;
            });

            return div;
        };
        legend.addTo(map);

        // Filter functionality
        let activeFilter = null;
        function filterMarkers(combo) {
            const legendItem = document.getElementById('legend-' + combo.replace(/[, ]/g, '_'));

            if (activeFilter === combo) {
                // Remove filter
                markers.clearLayers();
                Object.values(markerLayers).flat().forEach(m => markers.addLayer(m));
                activeFilter = null;
                legendItem.classList.remove('filtered');
            } else {
                // Apply filter
                markers.clearLayers();
                if (markerLayers[combo]) {
                    markerLayers[combo].forEach(m => markers.addLayer(m));
                }
                // Remove filtered class from all items
                document.querySelectorAll('.legend-item').forEach(item => {
                    item.classList.remove('filtered');
                });
                legendItem.classList.add('filtered');
                activeFilter = combo;
            }
        }

        // Add search functionality
        const searchBox = L.control({position: 'topright'});
        searchBox.onAdd = function(map) {
            const div = L.DomUtil.create('div', 'search-box');
            div.innerHTML = `
                <input type="text" id="searchInput" placeholder="Search store #, city, state...">
            `;
            // Prevent map clicks when interacting with search box
            L.DomEvent.disableClickPropagation(div);
            return div;
        };
        searchBox.addTo(map);

        // Search handler
        setTimeout(() => {
            const searchInput = document.getElementById('searchInput');
            if (searchInput) {
                searchInput.addEventListener('input', (e) => {
                    const searchTerm = e.target.value.toLowerCase();
                    if (searchTerm.length < 2) return;

                    const matching = STORE_DATA.features.find(f => {
                        const props = f.properties;
                        return props.store_number.toString().includes(searchTerm) ||
                               props.city.toLowerCase().includes(searchTerm) ||
                               props.state.toLowerCase().includes(searchTerm);
                    });

                    if (matching) {
                        const coords = matching.geometry.coordinates;
                        map.setView([coords[1], coords[0]], 12);
                    }
                });
            }
        }, 100);
    </script>
</body>
</html>'''

    return html


def generate_combined_html(geojson, state_summary):
    """Generate a single HTML file with both map and table views"""

    html = '''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Catalyst Pet Store Distribution</title>

    <!-- Leaflet CSS -->
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css" />
    <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css" />

    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f5f5; }

        #header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px 30px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }

        #header h1 {
            font-size: 28px;
            margin-bottom: 8px;
        }

        #header p {
            font-size: 14px;
            opacity: 0.95;
        }

        .nav-links {
            margin-top: 12px;
        }

        .nav-links button {
            color: white;
            background: transparent;
            text-decoration: none;
            margin-right: 15px;
            padding: 8px 16px;
            border: 1px solid rgba(255,255,255,0.5);
            border-radius: 4px;
            transition: all 0.3s;
            display: inline-block;
            font-size: 14px;
            cursor: pointer;
            font-family: inherit;
        }

        .nav-links button:hover {
            background: white;
            color: #667eea;
        }

        .nav-links button.active {
            background: rgba(255,255,255,0.2);
            border-color: white;
        }

        .view-container {
            display: none;
        }

        .view-container.active {
            display: block;
        }

        #map {
            height: calc(100vh - 140px);
            width: 100%;
        }

        .legend {
            background: white;
            padding: 12px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
            max-height: 500px;
            overflow-y: auto;
            min-width: 200px;
        }

        .legend h4 {
            margin: 0 0 10px 0;
            font-size: 14px;
            color: #333;
            border-bottom: 2px solid #667eea;
            padding-bottom: 6px;
        }

        .legend-item {
            display: flex;
            align-items: center;
            margin: 6px 0;
            cursor: pointer;
            padding: 6px;
            border-radius: 4px;
            transition: background 0.2s;
            font-size: 12px;
        }

        .legend-item:hover {
            background: #f5f5f5;
        }

        .legend-item.filtered {
            background: #e3f2fd;
            border: 1px solid #2196F3;
        }

        .legend-color {
            width: 18px;
            height: 18px;
            border-radius: 50%;
            margin-right: 8px;
            border: 2px solid #333;
            flex-shrink: 0;
        }

        .legend-text {
            font-size: 11px;
            line-height: 1.3;
        }

        .search-box {
            position: absolute;
            top: 10px;
            right: 10px;
            z-index: 1000;
            background: white;
            padding: 10px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
        }

        .search-box input {
            width: 250px;
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 13px;
        }

        .search-box input:focus {
            outline: none;
            border-color: #667eea;
        }

        .leaflet-popup-content {
            font-size: 13px;
            line-height: 1.5;
        }

        .leaflet-popup-content strong {
            color: #667eea;
        }

        .container {
            max-width: 1200px;
            margin: 30px auto;
            padding: 0 20px;
        }

        .search-filter {
            margin: 0 0 20px 0;
            padding: 15px;
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }

        .search-filter input {
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 4px;
            width: 100%;
            max-width: 400px;
            font-size: 14px;
        }

        .search-filter input:focus {
            outline: none;
            border-color: #667eea;
        }

        .table-wrapper {
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            overflow: hidden;
        }

        table {
            width: 100%;
            border-collapse: collapse;
        }

        thead {
            background: #667eea;
            color: white;
        }

        th {
            padding: 15px 12px;
            text-align: left;
            font-weight: 600;
            cursor: pointer;
            user-select: none;
            font-size: 13px;
        }

        th:hover {
            background: #5568d3;
        }

        th::after {
            content: ' ⇅';
            opacity: 0.5;
            font-size: 11px;
        }

        th.number-header {
            text-align: right;
        }

        td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #f0f0f0;
            font-size: 13px;
        }

        tbody tr:hover {
            background: #f8f9fa;
        }

        tfoot {
            font-weight: bold;
            background: #f5f5f5;
            border-top: 2px solid #667eea;
        }

        tfoot td {
            padding: 15px 12px;
            border-bottom: none;
        }

        .number-cell {
            text-align: right;
            font-family: 'Courier New', monospace;
        }

        @media (max-width: 768px) {
            #map { height: calc(100vh - 160px); }
            .search-box input { width: 180px; }
            .legend { max-height: 300px; }
            .container { padding: 0 10px; }
            th, td { padding: 10px 8px; font-size: 12px; }
            .search-filter { padding: 10px; }
        }
    </style>
</head>
<body>
    <div id="header">
        <h1>Catalyst Pet - Walmart Store Distribution</h1>
        <p>''' + f"{geojson['metadata']['geocoded_stores']:,}" + ''' stores across 50 states | 4 SKUs | 11 unique combinations</p>
        <div class="nav-links">
            <button id="mapViewBtn" class="active" onclick="switchView('map')">Map View</button>
            <button id="tableViewBtn" onclick="switchView('table')">Data Table</button>
        </div>
    </div>

    <!-- Map View -->
    <div id="mapView" class="view-container active">
        <div id="map"></div>
    </div>

    <!-- Table View -->
    <div id="tableView" class="view-container">
        <div class="container">
            <div class="search-filter">
                <input type="text" id="stateFilter" placeholder="Filter by state name...">
            </div>

            <div class="table-wrapper">
                <table id="dataTable">
                    <thead>
                        <tr>
                            <th onclick="sortTable(0)">State</th>
                            <th onclick="sortTable(1)" class="number-header">Total Stores</th>
                            <th onclick="sortTable(2)" class="number-header">15O</th>
                            <th onclick="sortTable(3)" class="number-header">15U</th>
                            <th onclick="sortTable(4)" class="number-header">34O</th>
                            <th onclick="sortTable(5)" class="number-header">34U</th>
                        </tr>
                    </thead>
                    <tbody id="tableBody">
                        <!-- Generated by JavaScript -->
                    </tbody>
                    <tfoot>
                        <tr>
                            <td>TOTAL</td>
                            <td class="number-cell" id="totalStores">-</td>
                            <td class="number-cell" id="total15Orig">-</td>
                            <td class="number-cell" id="total15Unscen">-</td>
                            <td class="number-cell" id="total34Orig">-</td>
                            <td class="number-cell" id="totalPet34">-</td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        </div>
    </div>

    <!-- Leaflet JS -->
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <script src="https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js"></script>

    <!-- Embedded Data -->
    <script>
    const STORE_DATA = ''' + json.dumps(geojson, separators=(',', ':')) + ''';
    const SKU_COLORS = ''' + json.dumps(SKU_COLORS, separators=(',', ':')) + ''';
    const STATE_SUMMARY = ''' + json.dumps(state_summary, separators=(',', ':')) + ''';
    </script>

    <script>
        // View switching
        function switchView(view) {
            if (view === 'map') {
                document.getElementById('mapView').classList.add('active');
                document.getElementById('tableView').classList.remove('active');
                document.getElementById('mapViewBtn').classList.add('active');
                document.getElementById('tableViewBtn').classList.remove('active');
                // Invalidate map size when switching back to map view
                if (window.map) {
                    setTimeout(() => map.invalidateSize(), 100);
                }
            } else {
                document.getElementById('mapView').classList.remove('active');
                document.getElementById('tableView').classList.add('active');
                document.getElementById('mapViewBtn').classList.remove('active');
                document.getElementById('tableViewBtn').classList.add('active');
            }
        }

        // Format SKU display with abbreviations
        function formatSkuDisplay(skuCombination) {
            const skuAbbrev = {
                'CATALYST15ORIG': '15O',
                'CATALYST15UNSCEN': '15U',
                'CATALYST34LBORIGINAL': '34O',
                'CATALYSTPET34LBUNSCE': '34U'
            };

            const skus = skuCombination.split(', ');
            const abbreviated = skus.map(sku => skuAbbrev[sku] || sku);
            return abbreviated.join(' / ');
        }

        // Initialize map
        const map = L.map('map').setView([39.8283, -98.5795], 5);

        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
            maxZoom: 18
        }).addTo(map);

        const markers = L.markerClusterGroup({
            maxClusterRadius: 50,
            spiderfyOnMaxZoom: true,
            showCoverageOnHover: false,
            disableClusteringAtZoom: 10
        });

        const markerLayers = {};

        STORE_DATA.features.forEach(feature => {
            const props = feature.properties;
            const coords = feature.geometry.coordinates;

            if (coords[1] && coords[0]) {
                const marker = L.circleMarker([coords[1], coords[0]], {
                    radius: 6,
                    fillColor: props.color,
                    color: '#000',
                    weight: 1,
                    opacity: 1,
                    fillOpacity: 0.8
                });

                marker.bindPopup(`
                    <strong>${props.store_name}</strong><br>
                    Store #${props.store_number}<br>
                    ${props.street_address}<br>
                    ${props.city}, ${props.state}<br>
                    <br>
                    <strong>SKUs (${props.sku_count}):</strong><br>
                    ${props.skus.map(sku => '• ' + sku).join('<br>')}
                `);

                if (!markerLayers[props.sku_combination]) {
                    markerLayers[props.sku_combination] = [];
                }
                markerLayers[props.sku_combination].push(marker);
                markers.addLayer(marker);
            }
        });

        map.addLayer(markers);

        // Legend
        const legend = L.control({position: 'bottomleft'});
        legend.onAdd = function(map) {
            const div = L.DomUtil.create('div', 'legend');
            div.innerHTML = '<h4>SKU Combinations</h4>';

            const comboCounts = {};
            STORE_DATA.features.forEach(f => {
                const combo = f.properties.sku_combination;
                comboCounts[combo] = (comboCounts[combo] || 0) + 1;
            });

            const sortedCombos = Object.entries(comboCounts).sort((a, b) => b[1] - a[1]);

            sortedCombos.forEach(([combo, count]) => {
                const color = SKU_COLORS[combo];
                const displayName = formatSkuDisplay(combo);

                div.innerHTML += `
                    <div class="legend-item" onclick="filterMarkers('${combo}')" id="legend-${combo.replace(/[, ]/g, '_')}">
                        <div class="legend-color" style="background-color: ${color}"></div>
                        <div class="legend-text"><strong>${count}</strong> - ${displayName}</div>
                    </div>
                `;
            });

            return div;
        };
        legend.addTo(map);

        // Filter functionality
        let activeFilter = null;
        function filterMarkers(combo) {
            const legendItem = document.getElementById('legend-' + combo.replace(/[, ]/g, '_'));

            if (activeFilter === combo) {
                markers.clearLayers();
                Object.values(markerLayers).flat().forEach(m => markers.addLayer(m));
                activeFilter = null;
                legendItem.classList.remove('filtered');
            } else {
                markers.clearLayers();
                if (markerLayers[combo]) {
                    markerLayers[combo].forEach(m => markers.addLayer(m));
                }
                document.querySelectorAll('.legend-item').forEach(item => {
                    item.classList.remove('filtered');
                });
                legendItem.classList.add('filtered');
                activeFilter = combo;
            }
        }

        // Search box
        const searchBox = L.control({position: 'topright'});
        searchBox.onAdd = function(map) {
            const div = L.DomUtil.create('div', 'search-box');
            div.innerHTML = `<input type="text" id="searchInput" placeholder="Search store #, city, state...">`;
            L.DomEvent.disableClickPropagation(div);
            return div;
        };
        searchBox.addTo(map);

        setTimeout(() => {
            const searchInput = document.getElementById('searchInput');
            if (searchInput) {
                searchInput.addEventListener('input', (e) => {
                    const searchTerm = e.target.value.toLowerCase();
                    if (searchTerm.length < 2) return;

                    const matching = STORE_DATA.features.find(f => {
                        const props = f.properties;
                        return props.store_number.toString().includes(searchTerm) ||
                               props.city.toLowerCase().includes(searchTerm) ||
                               props.state.toLowerCase().includes(searchTerm);
                    });

                    if (matching) {
                        const coords = matching.geometry.coordinates;
                        map.setView([coords[1], coords[0]], 12);
                    }
                });
            }
        }, 100);

        // Table functionality
        let currentSort = {column: -1, ascending: false};
        let currentData = STATE_SUMMARY.states;

        function populateTable(data) {
            const tbody = document.getElementById('tableBody');
            tbody.innerHTML = '';

            data.forEach(state => {
                const row = tbody.insertRow();
                row.innerHTML = `
                    <td>${state.state}</td>
                    <td class="number-cell">${state.total_stores.toLocaleString()}</td>
                    <td class="number-cell">${state.sku_counts.CATALYST15ORIG.toLocaleString()}</td>
                    <td class="number-cell">${state.sku_counts.CATALYST15UNSCEN.toLocaleString()}</td>
                    <td class="number-cell">${state.sku_counts.CATALYST34LBORIGINAL.toLocaleString()}</td>
                    <td class="number-cell">${state.sku_counts.CATALYSTPET34LBUNSCE.toLocaleString()}</td>
                `;
            });

            updateTotals(data);
        }

        function updateTotals(data) {
            const totals = STATE_SUMMARY.totals;
            document.getElementById('totalStores').textContent = totals.total_stores.toLocaleString();
            document.getElementById('total15Orig').textContent = totals.CATALYST15ORIG.toLocaleString();
            document.getElementById('total15Unscen').textContent = totals.CATALYST15UNSCEN.toLocaleString();
            document.getElementById('total34Orig').textContent = totals.CATALYST34LBORIGINAL.toLocaleString();
            document.getElementById('totalPet34').textContent = totals.CATALYSTPET34LBUNSCE.toLocaleString();
        }

        function sortTable(columnIndex) {
            const data = [...currentData];

            if (currentSort.column === columnIndex) {
                currentSort.ascending = !currentSort.ascending;
            } else {
                currentSort.column = columnIndex;
                currentSort.ascending = false;
            }

            data.sort((a, b) => {
                let aVal, bVal;

                switch(columnIndex) {
                    case 0: aVal = a.state; bVal = b.state; break;
                    case 1: aVal = a.total_stores; bVal = b.total_stores; break;
                    case 2: aVal = a.sku_counts.CATALYST15ORIG; bVal = b.sku_counts.CATALYST15ORIG; break;
                    case 3: aVal = a.sku_counts.CATALYST15UNSCEN; bVal = b.sku_counts.CATALYST15UNSCEN; break;
                    case 4: aVal = a.sku_counts.CATALYST34LBORIGINAL; bVal = b.sku_counts.CATALYST34LBORIGINAL; break;
                    case 5: aVal = a.sku_counts.CATALYSTPET34LBUNSCE; bVal = b.sku_counts.CATALYSTPET34LBUNSCE; break;
                }

                if (typeof aVal === 'string') {
                    return currentSort.ascending ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
                } else {
                    return currentSort.ascending ? aVal - bVal : bVal - aVal;
                }
            });

            currentData = data;
            populateTable(data);
        }

        document.getElementById('stateFilter').addEventListener('input', (e) => {
            const filter = e.target.value.toLowerCase();
            const filtered = STATE_SUMMARY.states.filter(s =>
                s.state.toLowerCase().includes(filter)
            );
            currentData = filtered;
            populateTable(filtered);
        });

        // Initialize table
        populateTable(currentData);
        sortTable(1);
    </script>
</body>
</html>'''

    return html


def generate_three_tab_html(geojson, state_summary):
    """Generate a single HTML with 3 tabs: SKU map, Simple map, Data table"""

    html = '''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Catalyst Pet Store Distribution</title>

    <!-- Leaflet CSS -->
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css" />
    <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css" />

    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f5f5; }

        #header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px 30px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }

        #header h1 {
            font-size: 28px;
            margin-bottom: 8px;
        }

        #header p {
            font-size: 14px;
            opacity: 0.95;
        }

        .nav-links {
            margin-top: 12px;
        }

        .nav-links button {
            color: white;
            background: transparent;
            text-decoration: none;
            margin-right: 15px;
            padding: 8px 16px;
            border: 1px solid rgba(255,255,255,0.5);
            border-radius: 4px;
            transition: all 0.3s;
            display: inline-block;
            font-size: 14px;
            cursor: pointer;
            font-family: inherit;
        }

        .nav-links button:hover {
            background: white;
            color: #667eea;
        }

        .nav-links button.active {
            background: rgba(255,255,255,0.2);
            border-color: white;
        }

        .view-container {
            display: none;
        }

        .view-container.active {
            display: block;
        }

        #mapSku, #mapSimple {
            height: calc(100vh - 140px);
            width: 100%;
        }

        .legend {
            background: white;
            padding: 12px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
            max-height: 500px;
            overflow-y: auto;
            min-width: 200px;
        }

        .legend h4 {
            margin: 0 0 10px 0;
            font-size: 14px;
            color: #333;
            border-bottom: 2px solid #667eea;
            padding-bottom: 6px;
        }

        .legend-item {
            display: flex;
            align-items: center;
            margin: 6px 0;
            cursor: pointer;
            padding: 6px;
            border-radius: 4px;
            transition: background 0.2s;
            font-size: 12px;
        }

        .legend-item:hover {
            background: #f5f5f5;
        }

        .legend-item.filtered {
            background: #e3f2fd;
            border: 1px solid #2196F3;
        }

        .legend-color {
            width: 18px;
            height: 18px;
            border-radius: 50%;
            margin-right: 8px;
            border: 2px solid #333;
            flex-shrink: 0;
        }

        .legend-text {
            font-size: 11px;
            line-height: 1.3;
        }

        .info-box {
            background: white;
            padding: 12px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
        }

        .info-box h4 {
            margin: 0 0 10px 0;
            font-size: 14px;
            color: #333;
            border-bottom: 2px solid #667eea;
            padding-bottom: 6px;
        }

        .info-box p {
            margin: 5px 0;
            font-size: 12px;
            color: #666;
        }

        .search-box {
            position: absolute;
            top: 10px;
            right: 10px;
            z-index: 1000;
            background: white;
            padding: 10px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
        }

        .search-box input {
            width: 250px;
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 13px;
        }

        .search-box input:focus {
            outline: none;
            border-color: #667eea;
        }

        .leaflet-popup-content {
            font-size: 13px;
            line-height: 1.5;
        }

        .leaflet-popup-content strong {
            color: #667eea;
        }

        .container {
            max-width: 1200px;
            margin: 30px auto;
            padding: 0 20px;
        }

        .search-filter {
            margin: 0 0 20px 0;
            padding: 15px;
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }

        .search-filter input {
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 4px;
            width: 100%;
            max-width: 400px;
            font-size: 14px;
        }

        .search-filter input:focus {
            outline: none;
            border-color: #667eea;
        }

        .table-wrapper {
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            overflow: hidden;
        }

        table {
            width: 100%;
            border-collapse: collapse;
        }

        thead {
            background: #667eea;
            color: white;
        }

        th {
            padding: 15px 12px;
            text-align: left;
            font-weight: 600;
            cursor: pointer;
            user-select: none;
            font-size: 13px;
        }

        th:hover {
            background: #5568d3;
        }

        th::after {
            content: ' ⇅';
            opacity: 0.5;
            font-size: 11px;
        }

        th.number-header {
            text-align: right;
        }

        td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #f0f0f0;
            font-size: 13px;
        }

        tbody tr:hover {
            background: #f8f9fa;
        }

        tfoot {
            font-weight: bold;
            background: #f5f5f5;
            border-top: 2px solid #667eea;
        }

        tfoot td {
            padding: 15px 12px;
            border-bottom: none;
        }

        .number-cell {
            text-align: right;
            font-family: 'Courier New', monospace;
        }

        @media (max-width: 768px) {
            #mapSku, #mapSimple { height: calc(100vh - 160px); }
            .search-box input { width: 180px; }
            .legend { max-height: 300px; }
            .container { padding: 0 10px; }
            th, td { padding: 10px 8px; font-size: 12px; }
            .search-filter { padding: 10px; }
        }
    </style>
</head>
<body>
    <div id="header">
        <h1>Catalyst Pet - Walmart Store Distribution</h1>
        <p>''' + f"{geojson['metadata']['geocoded_stores']:,}" + ''' stores across 50 states | 4 SKUs</p>
        <div class="nav-links">
            <button id="skuMapBtn" class="active" onclick="switchView('skuMap')">SKU Map</button>
            <button id="simpleMapBtn" onclick="switchView('simpleMap')">Simple Map</button>
            <button id="tableBtn" onclick="switchView('table')">Data Table</button>
        </div>
    </div>

    <!-- SKU Color-Coded Map View -->
    <div id="skuMapView" class="view-container active">
        <div id="mapSku"></div>
    </div>

    <!-- Simple Map View -->
    <div id="simpleMapView" class="view-container">
        <div id="mapSimple"></div>
    </div>

    <!-- Table View -->
    <div id="tableView" class="view-container">
        <div class="container">
            <div class="search-filter">
                <input type="text" id="stateFilter" placeholder="Filter by state name...">
            </div>

            <div class="table-wrapper">
                <table id="dataTable">
                    <thead>
                        <tr>
                            <th onclick="sortTable(0)">State</th>
                            <th onclick="sortTable(1)" class="number-header">Total Stores</th>
                            <th onclick="sortTable(2)" class="number-header">15O</th>
                            <th onclick="sortTable(3)" class="number-header">15U</th>
                            <th onclick="sortTable(4)" class="number-header">34O</th>
                            <th onclick="sortTable(5)" class="number-header">34U</th>
                        </tr>
                    </thead>
                    <tbody id="tableBody">
                        <!-- Generated by JavaScript -->
                    </tbody>
                    <tfoot>
                        <tr>
                            <td>TOTAL</td>
                            <td class="number-cell" id="totalStores">-</td>
                            <td class="number-cell" id="total15Orig">-</td>
                            <td class="number-cell" id="total15Unscen">-</td>
                            <td class="number-cell" id="total34Orig">-</td>
                            <td class="number-cell" id="totalPet34">-</td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        </div>
    </div>

    <!-- Leaflet JS -->
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <script src="https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js"></script>

    <!-- Embedded Data -->
    <script>
    const STORE_DATA = ''' + json.dumps(geojson, separators=(',', ':')) + ''';
    const SKU_COLORS = ''' + json.dumps(SKU_COLORS, separators=(',', ':')) + ''';
    const STATE_SUMMARY = ''' + json.dumps(state_summary, separators=(',', ':')) + ''';

    let mapSku, mapSimple;
    let currentView = 'skuMap';

    // View switching
    function switchView(view) {
        // Hide all views
        document.getElementById('skuMapView').classList.remove('active');
        document.getElementById('simpleMapView').classList.remove('active');
        document.getElementById('tableView').classList.remove('active');

        // Deactivate all buttons
        document.getElementById('skuMapBtn').classList.remove('active');
        document.getElementById('simpleMapBtn').classList.remove('active');
        document.getElementById('tableBtn').classList.remove('active');

        // Show selected view
        if (view === 'skuMap') {
            document.getElementById('skuMapView').classList.add('active');
            document.getElementById('skuMapBtn').classList.add('active');
            if (!mapSku) initSkuMap();
            else setTimeout(() => mapSku.invalidateSize(), 100);
        } else if (view === 'simpleMap') {
            document.getElementById('simpleMapView').classList.add('active');
            document.getElementById('simpleMapBtn').classList.add('active');
            if (!mapSimple) initSimpleMap();
            else setTimeout(() => mapSimple.invalidateSize(), 100);
        } else {
            document.getElementById('tableView').classList.add('active');
            document.getElementById('tableBtn').classList.add('active');
        }

        currentView = view;
    }

    // Format SKU display
    function formatSkuDisplay(skuCombination) {
        const skuAbbrev = {
            'CATALYST15ORIG': '15O',
            'CATALYST15UNSCEN': '15U',
            'CATALYST34LBORIGINAL': '34O',
            'CATALYSTPET34LBUNSCE': '34U'
        };
        const skus = skuCombination.split(', ');
        const abbreviated = skus.map(sku => skuAbbrev[sku] || sku);
        return abbreviated.join(' / ');
    }

    // Initialize SKU Color-Coded Map
    function initSkuMap() {
        mapSku = L.map('mapSku').setView([39.8283, -98.5795], 5);

        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
            maxZoom: 18
        }).addTo(mapSku);

        const markers = L.markerClusterGroup({
            maxClusterRadius: 50,
            spiderfyOnMaxZoom: true,
            showCoverageOnHover: false,
            disableClusteringAtZoom: 10
        });

        const markerLayers = {};

        STORE_DATA.features.forEach(feature => {
            const props = feature.properties;
            const coords = feature.geometry.coordinates;

            if (coords[1] && coords[0]) {
                const marker = L.circleMarker([coords[1], coords[0]], {
                    radius: 6,
                    fillColor: props.color,
                    color: '#000',
                    weight: 1,
                    opacity: 1,
                    fillOpacity: 0.8
                });

                marker.bindPopup(`
                    <strong>${props.store_name}</strong><br>
                    Store #${props.store_number}<br>
                    ${props.street_address}<br>
                    ${props.city}, ${props.state}<br>
                    <br>
                    <strong>SKUs (${props.sku_count}):</strong><br>
                    ${props.skus.map(sku => '• ' + sku).join('<br>')}
                `);

                if (!markerLayers[props.sku_combination]) {
                    markerLayers[props.sku_combination] = [];
                }
                markerLayers[props.sku_combination].push(marker);
                markers.addLayer(marker);
            }
        });

        mapSku.addLayer(markers);

        // Legend
        const legend = L.control({position: 'bottomleft'});
        legend.onAdd = function(map) {
            const div = L.DomUtil.create('div', 'legend');
            div.innerHTML = '<h4>SKU Combinations</h4>';

            const comboCounts = {};
            STORE_DATA.features.forEach(f => {
                const combo = f.properties.sku_combination;
                comboCounts[combo] = (comboCounts[combo] || 0) + 1;
            });

            const sortedCombos = Object.entries(comboCounts).sort((a, b) => b[1] - a[1]);

            sortedCombos.forEach(([combo, count]) => {
                const color = SKU_COLORS[combo];
                const displayName = formatSkuDisplay(combo);

                div.innerHTML += `
                    <div class="legend-item" onclick="filterMarkers('${combo.replace(/'/g, "\\\'")}')" id="legend-${combo.replace(/[, ]/g, '_')}">
                        <div class="legend-color" style="background-color: ${color}"></div>
                        <div class="legend-text"><strong>${count}</strong> - ${displayName}</div>
                    </div>
                `;
            });

            return div;
        };
        legend.addTo(mapSku);

        // Search box
        const searchBox = L.control({position: 'topright'});
        searchBox.onAdd = function(map) {
            const div = L.DomUtil.create('div', 'search-box');
            div.innerHTML = `<input type="text" id="searchInputSku" placeholder="Search store #, city, state...">`;
            L.DomEvent.disableClickPropagation(div);
            return div;
        };
        searchBox.addTo(mapSku);

        setTimeout(() => {
            const searchInput = document.getElementById('searchInputSku');
            if (searchInput) {
                searchInput.addEventListener('input', (e) => {
                    const searchTerm = e.target.value.toLowerCase();
                    if (searchTerm.length < 2) return;

                    const matching = STORE_DATA.features.find(f => {
                        const props = f.properties;
                        return props.store_number.toString().includes(searchTerm) ||
                               props.city.toLowerCase().includes(searchTerm) ||
                               props.state.toLowerCase().includes(searchTerm);
                    });

                    if (matching) {
                        const coords = matching.geometry.coordinates;
                        mapSku.setView([coords[1], coords[0]], 12);
                    }
                });
            }
        }, 100);

        // Store references for filtering
        window.skuMarkers = markers;
        window.skuMarkerLayers = markerLayers;
    }

    // Filter markers
    let activeFilter = null;
    function filterMarkers(combo) {
        const markers = window.skuMarkers;
        const markerLayers = window.skuMarkerLayers;
        const legendItem = document.getElementById('legend-' + combo.replace(/[, ]/g, '_'));

        if (activeFilter === combo) {
            markers.clearLayers();
            Object.values(markerLayers).flat().forEach(m => markers.addLayer(m));
            activeFilter = null;
            legendItem.classList.remove('filtered');
        } else {
            markers.clearLayers();
            if (markerLayers[combo]) {
                markerLayers[combo].forEach(m => markers.addLayer(m));
            }
            document.querySelectorAll('.legend-item').forEach(item => {
                item.classList.remove('filtered');
            });
            legendItem.classList.add('filtered');
            activeFilter = combo;
        }
    }

    // Initialize Simple Map
    function initSimpleMap() {
        mapSimple = L.map('mapSimple').setView([39.8283, -98.5795], 5);

        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
            maxZoom: 18
        }).addTo(mapSimple);

        const markers = L.markerClusterGroup({
            maxClusterRadius: 50,
            spiderfyOnMaxZoom: true,
            showCoverageOnHover: false,
            disableClusteringAtZoom: 10
        });

        STORE_DATA.features.forEach(feature => {
            const props = feature.properties;
            const coords = feature.geometry.coordinates;

            if (coords[1] && coords[0]) {
                const marker = L.circleMarker([coords[1], coords[0]], {
                    radius: 6,
                    fillColor: '#4CAF50',
                    color: '#000',
                    weight: 1,
                    opacity: 1,
                    fillOpacity: 0.8
                });

                marker.bindPopup(`
                    <strong>${props.store_name}</strong><br>
                    Store #${props.store_number}<br>
                    ${props.street_address}<br>
                    ${props.city}, ${props.state}<br>
                    <br>
                    <strong>SKUs (${props.sku_count}):</strong><br>
                    ${props.skus.map(sku => '• ' + sku).join('<br>')}
                `);

                markers.addLayer(marker);
            }
        });

        mapSimple.addLayer(markers);

        // Info box
        const info = L.control({position: 'bottomleft'});
        info.onAdd = function(map) {
            const div = L.DomUtil.create('div', 'info-box');
            div.innerHTML = `
                <h4>Store Locations</h4>
                <p><strong>${STORE_DATA.features.length.toLocaleString()}</strong> Walmart stores</p>
                <p>Carrying Catalyst Pet products</p>
            `;
            return div;
        };
        info.addTo(mapSimple);

        // Search box
        const searchBox = L.control({position: 'topright'});
        searchBox.onAdd = function(map) {
            const div = L.DomUtil.create('div', 'search-box');
            div.innerHTML = `<input type="text" id="searchInputSimple" placeholder="Search store #, city, state...">`;
            L.DomEvent.disableClickPropagation(div);
            return div;
        };
        searchBox.addTo(mapSimple);

        setTimeout(() => {
            const searchInput = document.getElementById('searchInputSimple');
            if (searchInput) {
                searchInput.addEventListener('input', (e) => {
                    const searchTerm = e.target.value.toLowerCase();
                    if (searchTerm.length < 2) return;

                    const matching = STORE_DATA.features.find(f => {
                        const props = f.properties;
                        return props.store_number.toString().includes(searchTerm) ||
                               props.city.toLowerCase().includes(searchTerm) ||
                               props.state.toLowerCase().includes(searchTerm);
                    });

                    if (matching) {
                        const coords = matching.geometry.coordinates;
                        mapSimple.setView([coords[1], coords[0]], 12);
                    }
                });
            }
        }, 100);
    }

    // Table functionality
    let currentSort = {column: -1, ascending: false};
    let currentData = STATE_SUMMARY.states;

    function populateTable(data) {
        const tbody = document.getElementById('tableBody');
        tbody.innerHTML = '';

        data.forEach(state => {
            const row = tbody.insertRow();
            row.innerHTML = `
                <td>${state.state}</td>
                <td class="number-cell">${state.total_stores.toLocaleString()}</td>
                <td class="number-cell">${state.sku_counts.CATALYST15ORIG.toLocaleString()}</td>
                <td class="number-cell">${state.sku_counts.CATALYST15UNSCEN.toLocaleString()}</td>
                <td class="number-cell">${state.sku_counts.CATALYST34LBORIGINAL.toLocaleString()}</td>
                <td class="number-cell">${state.sku_counts.CATALYSTPET34LBUNSCE.toLocaleString()}</td>
            `;
        });

        updateTotals(data);
    }

    function updateTotals(data) {
        const totals = STATE_SUMMARY.totals;
        document.getElementById('totalStores').textContent = totals.total_stores.toLocaleString();
        document.getElementById('total15Orig').textContent = totals.CATALYST15ORIG.toLocaleString();
        document.getElementById('total15Unscen').textContent = totals.CATALYST15UNSCEN.toLocaleString();
        document.getElementById('total34Orig').textContent = totals.CATALYST34LBORIGINAL.toLocaleString();
        document.getElementById('totalPet34').textContent = totals.CATALYSTPET34LBUNSCE.toLocaleString();
    }

    function sortTable(columnIndex) {
        const data = [...currentData];

        if (currentSort.column === columnIndex) {
            currentSort.ascending = !currentSort.ascending;
        } else {
            currentSort.column = columnIndex;
            currentSort.ascending = false;
        }

        data.sort((a, b) => {
            let aVal, bVal;

            switch(columnIndex) {
                case 0: aVal = a.state; bVal = b.state; break;
                case 1: aVal = a.total_stores; bVal = b.total_stores; break;
                case 2: aVal = a.sku_counts.CATALYST15ORIG; bVal = b.sku_counts.CATALYST15ORIG; break;
                case 3: aVal = a.sku_counts.CATALYST15UNSCEN; bVal = b.sku_counts.CATALYST15UNSCEN; break;
                case 4: aVal = a.sku_counts.CATALYST34LBORIGINAL; bVal = b.sku_counts.CATALYST34LBORIGINAL; break;
                case 5: aVal = a.sku_counts.CATALYSTPET34LBUNSCE; bVal = b.sku_counts.CATALYSTPET34LBUNSCE; break;
            }

            if (typeof aVal === 'string') {
                return currentSort.ascending ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
            } else {
                return currentSort.ascending ? aVal - bVal : bVal - aVal;
            }
        });

        currentData = data;
        populateTable(data);
    }

    document.getElementById('stateFilter').addEventListener('input', (e) => {
        const filter = e.target.value.toLowerCase();
        const filtered = STATE_SUMMARY.states.filter(s =>
            s.state.toLowerCase().includes(filter)
        );
        currentData = filtered;
        populateTable(filtered);
    });

    // Initialize
    initSkuMap();
    populateTable(currentData);
    sortTable(1);
    </script>
</body>
</html>'''

    return html


def generate_simple_map_html(geojson, state_summary):
    """Generate a simple HTML with all stores using the same marker color"""

    html = '''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Catalyst Pet Store Locations</title>

    <!-- Leaflet CSS -->
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css" />
    <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css" />

    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f5f5; }

        #header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px 30px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }

        #header h1 {
            font-size: 28px;
            margin-bottom: 8px;
        }

        #header p {
            font-size: 14px;
            opacity: 0.95;
        }

        .nav-links {
            margin-top: 12px;
        }

        .nav-links button {
            color: white;
            background: transparent;
            text-decoration: none;
            margin-right: 15px;
            padding: 8px 16px;
            border: 1px solid rgba(255,255,255,0.5);
            border-radius: 4px;
            transition: all 0.3s;
            display: inline-block;
            font-size: 14px;
            cursor: pointer;
            font-family: inherit;
        }

        .nav-links button:hover {
            background: white;
            color: #667eea;
        }

        .nav-links button.active {
            background: rgba(255,255,255,0.2);
            border-color: white;
        }

        .view-container {
            display: none;
        }

        .view-container.active {
            display: block;
        }

        #map {
            height: calc(100vh - 140px);
            width: 100%;
        }

        .info-box {
            background: white;
            padding: 12px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
        }

        .info-box h4 {
            margin: 0 0 10px 0;
            font-size: 14px;
            color: #333;
            border-bottom: 2px solid #667eea;
            padding-bottom: 6px;
        }

        .info-box p {
            margin: 5px 0;
            font-size: 12px;
            color: #666;
        }

        .search-box {
            position: absolute;
            top: 10px;
            right: 10px;
            z-index: 1000;
            background: white;
            padding: 10px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
        }

        .search-box input {
            width: 250px;
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 13px;
        }

        .search-box input:focus {
            outline: none;
            border-color: #667eea;
        }

        .leaflet-popup-content {
            font-size: 13px;
            line-height: 1.5;
        }

        .leaflet-popup-content strong {
            color: #667eea;
        }

        .container {
            max-width: 1200px;
            margin: 30px auto;
            padding: 0 20px;
        }

        .search-filter {
            margin: 0 0 20px 0;
            padding: 15px;
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }

        .search-filter input {
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 4px;
            width: 100%;
            max-width: 400px;
            font-size: 14px;
        }

        .search-filter input:focus {
            outline: none;
            border-color: #667eea;
        }

        .table-wrapper {
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            overflow: hidden;
        }

        table {
            width: 100%;
            border-collapse: collapse;
        }

        thead {
            background: #667eea;
            color: white;
        }

        th {
            padding: 15px 12px;
            text-align: left;
            font-weight: 600;
            cursor: pointer;
            user-select: none;
            font-size: 13px;
        }

        th:hover {
            background: #5568d3;
        }

        th::after {
            content: ' ⇅';
            opacity: 0.5;
            font-size: 11px;
        }

        th.number-header {
            text-align: right;
        }

        td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #f0f0f0;
            font-size: 13px;
        }

        tbody tr:hover {
            background: #f8f9fa;
        }

        tfoot {
            font-weight: bold;
            background: #f5f5f5;
            border-top: 2px solid #667eea;
        }

        tfoot td {
            padding: 15px 12px;
            border-bottom: none;
        }

        .number-cell {
            text-align: right;
            font-family: 'Courier New', monospace;
        }

        @media (max-width: 768px) {
            #map { height: calc(100vh - 160px); }
            .search-box input { width: 180px; }
            .container { padding: 0 10px; }
            th, td { padding: 10px 8px; font-size: 12px; }
            .search-filter { padding: 10px; }
        }
    </style>
</head>
<body>
    <div id="header">
        <h1>Catalyst Pet - Walmart Store Locations</h1>
        <p>''' + f"{geojson['metadata']['geocoded_stores']:,}" + ''' stores across 50 states</p>
        <div class="nav-links">
            <button id="mapViewBtn" class="active" onclick="switchView('map')">Map View</button>
            <button id="tableViewBtn" onclick="switchView('table')">Data Table</button>
        </div>
    </div>

    <!-- Map View -->
    <div id="mapView" class="view-container active">
        <div id="map"></div>
    </div>

    <!-- Table View -->
    <div id="tableView" class="view-container">
        <div class="container">
            <div class="search-filter">
                <input type="text" id="stateFilter" placeholder="Filter by state name...">
            </div>

            <div class="table-wrapper">
                <table id="dataTable">
                    <thead>
                        <tr>
                            <th onclick="sortTable(0)">State</th>
                            <th onclick="sortTable(1)" class="number-header">Total Stores</th>
                            <th onclick="sortTable(2)" class="number-header">15O</th>
                            <th onclick="sortTable(3)" class="number-header">15U</th>
                            <th onclick="sortTable(4)" class="number-header">34O</th>
                            <th onclick="sortTable(5)" class="number-header">34U</th>
                        </tr>
                    </thead>
                    <tbody id="tableBody">
                        <!-- Generated by JavaScript -->
                    </tbody>
                    <tfoot>
                        <tr>
                            <td>TOTAL</td>
                            <td class="number-cell" id="totalStores">-</td>
                            <td class="number-cell" id="total15Orig">-</td>
                            <td class="number-cell" id="total15Unscen">-</td>
                            <td class="number-cell" id="total34Orig">-</td>
                            <td class="number-cell" id="totalPet34">-</td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        </div>
    </div>

    <!-- Leaflet JS -->
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <script src="https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js"></script>

    <!-- Embedded Data -->
    <script>
    const STORE_DATA = ''' + json.dumps(geojson, separators=(',', ':')) + ''';
    const STATE_SUMMARY = ''' + json.dumps(state_summary, separators=(',', ':')) + ''';
    </script>

    <script>
        // View switching
        function switchView(view) {
            if (view === 'map') {
                document.getElementById('mapView').classList.add('active');
                document.getElementById('tableView').classList.remove('active');
                document.getElementById('mapViewBtn').classList.add('active');
                document.getElementById('tableViewBtn').classList.remove('active');
                if (window.map) {
                    setTimeout(() => map.invalidateSize(), 100);
                }
            } else {
                document.getElementById('mapView').classList.remove('active');
                document.getElementById('tableView').classList.add('active');
                document.getElementById('mapViewBtn').classList.remove('active');
                document.getElementById('tableViewBtn').classList.add('active');
            }
        }

        // Initialize map
        const map = L.map('map').setView([39.8283, -98.5795], 5);

        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
            maxZoom: 18
        }).addTo(map);

        const markers = L.markerClusterGroup({
            maxClusterRadius: 50,
            spiderfyOnMaxZoom: true,
            showCoverageOnHover: false,
            disableClusteringAtZoom: 10
        });

        // Add all markers with same color
        STORE_DATA.features.forEach(feature => {
            const props = feature.properties;
            const coords = feature.geometry.coordinates;

            if (coords[1] && coords[0]) {
                const marker = L.circleMarker([coords[1], coords[0]], {
                    radius: 6,
                    fillColor: '#4CAF50',  // Single green color for all stores
                    color: '#000',
                    weight: 1,
                    opacity: 1,
                    fillOpacity: 0.8
                });

                marker.bindPopup(`
                    <strong>${props.store_name}</strong><br>
                    Store #${props.store_number}<br>
                    ${props.street_address}<br>
                    ${props.city}, ${props.state}<br>
                    <br>
                    <strong>SKUs (${props.sku_count}):</strong><br>
                    ${props.skus.map(sku => '• ' + sku).join('<br>')}
                `);

                markers.addLayer(marker);
            }
        });

        map.addLayer(markers);

        // Add info box
        const info = L.control({position: 'bottomleft'});
        info.onAdd = function(map) {
            const div = L.DomUtil.create('div', 'info-box');
            div.innerHTML = `
                <h4>Store Locations</h4>
                <p><strong>${STORE_DATA.features.length.toLocaleString()}</strong> Walmart stores</p>
                <p>Carrying Catalyst Pet products</p>
            `;
            return div;
        };
        info.addTo(map);

        // Search box
        const searchBox = L.control({position: 'topright'});
        searchBox.onAdd = function(map) {
            const div = L.DomUtil.create('div', 'search-box');
            div.innerHTML = `<input type="text" id="searchInput" placeholder="Search store #, city, state...">`;
            L.DomEvent.disableClickPropagation(div);
            return div;
        };
        searchBox.addTo(map);

        setTimeout(() => {
            const searchInput = document.getElementById('searchInput');
            if (searchInput) {
                searchInput.addEventListener('input', (e) => {
                    const searchTerm = e.target.value.toLowerCase();
                    if (searchTerm.length < 2) return;

                    const matching = STORE_DATA.features.find(f => {
                        const props = f.properties;
                        return props.store_number.toString().includes(searchTerm) ||
                               props.city.toLowerCase().includes(searchTerm) ||
                               props.state.toLowerCase().includes(searchTerm);
                    });

                    if (matching) {
                        const coords = matching.geometry.coordinates;
                        map.setView([coords[1], coords[0]], 12);
                    }
                });
            }
        }, 100);

        // Table functionality
        let currentSort = {column: -1, ascending: false};
        let currentData = STATE_SUMMARY.states;

        function populateTable(data) {
            const tbody = document.getElementById('tableBody');
            tbody.innerHTML = '';

            data.forEach(state => {
                const row = tbody.insertRow();
                row.innerHTML = `
                    <td>${state.state}</td>
                    <td class="number-cell">${state.total_stores.toLocaleString()}</td>
                    <td class="number-cell">${state.sku_counts.CATALYST15ORIG.toLocaleString()}</td>
                    <td class="number-cell">${state.sku_counts.CATALYST15UNSCEN.toLocaleString()}</td>
                    <td class="number-cell">${state.sku_counts.CATALYST34LBORIGINAL.toLocaleString()}</td>
                    <td class="number-cell">${state.sku_counts.CATALYSTPET34LBUNSCE.toLocaleString()}</td>
                `;
            });

            updateTotals(data);
        }

        function updateTotals(data) {
            const totals = STATE_SUMMARY.totals;
            document.getElementById('totalStores').textContent = totals.total_stores.toLocaleString();
            document.getElementById('total15Orig').textContent = totals.CATALYST15ORIG.toLocaleString();
            document.getElementById('total15Unscen').textContent = totals.CATALYST15UNSCEN.toLocaleString();
            document.getElementById('total34Orig').textContent = totals.CATALYST34LBORIGINAL.toLocaleString();
            document.getElementById('totalPet34').textContent = totals.CATALYSTPET34LBUNSCE.toLocaleString();
        }

        function sortTable(columnIndex) {
            const data = [...currentData];

            if (currentSort.column === columnIndex) {
                currentSort.ascending = !currentSort.ascending;
            } else {
                currentSort.column = columnIndex;
                currentSort.ascending = false;
            }

            data.sort((a, b) => {
                let aVal, bVal;

                switch(columnIndex) {
                    case 0: aVal = a.state; bVal = b.state; break;
                    case 1: aVal = a.total_stores; bVal = b.total_stores; break;
                    case 2: aVal = a.sku_counts.CATALYST15ORIG; bVal = b.sku_counts.CATALYST15ORIG; break;
                    case 3: aVal = a.sku_counts.CATALYST15UNSCEN; bVal = b.sku_counts.CATALYST15UNSCEN; break;
                    case 4: aVal = a.sku_counts.CATALYST34LBORIGINAL; bVal = b.sku_counts.CATALYST34LBORIGINAL; break;
                    case 5: aVal = a.sku_counts.CATALYSTPET34LBUNSCE; bVal = b.sku_counts.CATALYSTPET34LBUNSCE; break;
                }

                if (typeof aVal === 'string') {
                    return currentSort.ascending ? aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
                } else {
                    return currentSort.ascending ? aVal - bVal : bVal - aVal;
                }
            });

            currentData = data;
            populateTable(data);
        }

        document.getElementById('stateFilter').addEventListener('input', (e) => {
            const filter = e.target.value.toLowerCase();
            const filtered = STATE_SUMMARY.states.filter(s =>
                s.state.toLowerCase().includes(filter)
            );
            currentData = filtered;
            populateTable(filtered);
        });

        // Initialize table
        populateTable(currentData);
        sortTable(1);
    </script>
</body>
</html>'''

    return html


def generate_table_html(state_summary):
    """Generate the data table HTML"""

    html = '''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Catalyst Pet - Store Data by State</title>

    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f5f5; }

        #header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px 30px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }

        #header h1 {
            font-size: 28px;
            margin-bottom: 8px;
        }

        #header p {
            font-size: 14px;
            opacity: 0.95;
        }

        .nav-links {
            margin-top: 12px;
        }

        .nav-links a {
            color: white;
            text-decoration: none;
            margin-right: 15px;
            padding: 8px 16px;
            border: 1px solid rgba(255,255,255,0.5);
            border-radius: 4px;
            transition: all 0.3s;
            display: inline-block;
            font-size: 14px;
        }

        .nav-links a:hover {
            background: white;
            color: #667eea;
        }

        .nav-links a.active {
            background: rgba(255,255,255,0.2);
            border-color: white;
        }

        .container {
            max-width: 1200px;
            margin: 30px auto;
            padding: 0 20px;
        }

        .search-filter {
            margin: 0 0 20px 0;
            padding: 15px;
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }

        .search-filter input {
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 4px;
            width: 100%;
            max-width: 400px;
            font-size: 14px;
        }

        .search-filter input:focus {
            outline: none;
            border-color: #667eea;
        }

        .table-wrapper {
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            overflow: hidden;
        }

        table {
            width: 100%;
            border-collapse: collapse;
        }

        thead {
            background: #667eea;
            color: white;
        }

        th {
            padding: 15px 12px;
            text-align: left;
            font-weight: 600;
            cursor: pointer;
            user-select: none;
            font-size: 13px;
        }

        th:hover {
            background: #5568d3;
        }

        th::after {
            content: ' ⇅';
            opacity: 0.5;
            font-size: 11px;
        }

        th.number-header {
            text-align: right;
        }

        td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #f0f0f0;
            font-size: 13px;
        }

        tbody tr:hover {
            background: #f8f9fa;
        }

        tfoot {
            font-weight: bold;
            background: #f5f5f5;
            border-top: 2px solid #667eea;
        }

        tfoot td {
            padding: 15px 12px;
            border-bottom: none;
        }

        .number-cell {
            text-align: right;
            font-family: 'Courier New', monospace;
        }

        @media (max-width: 768px) {
            .container { padding: 0 10px; }
            th, td { padding: 10px 8px; font-size: 12px; }
            .search-filter { padding: 10px; }
        }
    </style>
</head>
<body>
    <div id="header">
        <h1>Catalyst Pet - Store Count by State</h1>
        <p>SKU distribution across all 50 states</p>
        <div class="nav-links">
            <a href="index.html">Map View</a>
            <a href="data-table.html" class="active">Data Table</a>
        </div>
    </div>

    <div class="container">
        <div class="search-filter">
            <input type="text" id="stateFilter" placeholder="Filter by state name...">
        </div>

        <div class="table-wrapper">
            <table id="dataTable">
                <thead>
                    <tr>
                        <th onclick="sortTable(0)">State</th>
                        <th onclick="sortTable(1)" class="number-header">Total Stores</th>
                        <th onclick="sortTable(2)" class="number-header">CATALYST15ORIG</th>
                        <th onclick="sortTable(3)" class="number-header">CATALYST15UNSCEN</th>
                        <th onclick="sortTable(4)" class="number-header">CATALYST34LBORIGINAL</th>
                        <th onclick="sortTable(5)" class="number-header">CATALYSTPET34LBUNSCE</th>
                    </tr>
                </thead>
                <tbody id="tableBody">
                    <!-- Generated by JavaScript -->
                </tbody>
                <tfoot>
                    <tr>
                        <td>TOTAL</td>
                        <td class="number-cell" id="totalStores">-</td>
                        <td class="number-cell" id="total15Orig">-</td>
                        <td class="number-cell" id="total15Unscen">-</td>
                        <td class="number-cell" id="total34Orig">-</td>
                        <td class="number-cell" id="totalPet34">-</td>
                    </tr>
                </tfoot>
            </table>
        </div>
    </div>

    <script>
    const STATE_SUMMARY = ''' + json.dumps(state_summary, separators=(',', ':')) + ''';

    let currentSort = {column: -1, ascending: false};
    let currentData = STATE_SUMMARY.states;

    // Populate table
    function populateTable(data) {
        const tbody = document.getElementById('tableBody');
        tbody.innerHTML = '';

        data.forEach(state => {
            const row = tbody.insertRow();
            row.innerHTML = `
                <td>${state.state}</td>
                <td class="number-cell">${state.total_stores.toLocaleString()}</td>
                <td class="number-cell">${state.sku_counts.CATALYST15ORIG.toLocaleString()}</td>
                <td class="number-cell">${state.sku_counts.CATALYST15UNSCEN.toLocaleString()}</td>
                <td class="number-cell">${state.sku_counts.CATALYST34LBORIGINAL.toLocaleString()}</td>
                <td class="number-cell">${state.sku_counts.CATALYSTPET34LBUNSCE.toLocaleString()}</td>
            `;
        });

        updateTotals(data);
    }

    function updateTotals(data) {
        const totals = STATE_SUMMARY.totals;
        document.getElementById('totalStores').textContent = totals.total_stores.toLocaleString();
        document.getElementById('total15Orig').textContent = totals.CATALYST15ORIG.toLocaleString();
        document.getElementById('total15Unscen').textContent = totals.CATALYST15UNSCEN.toLocaleString();
        document.getElementById('total34Orig').textContent = totals.CATALYST34LBORIGINAL.toLocaleString();
        document.getElementById('totalPet34').textContent = totals.CATALYSTPET34LBUNSCE.toLocaleString();
    }

    // Sort table
    function sortTable(columnIndex) {
        const data = [...currentData];

        if (currentSort.column === columnIndex) {
            currentSort.ascending = !currentSort.ascending;
        } else {
            currentSort.column = columnIndex;
            currentSort.ascending = false;
        }

        data.sort((a, b) => {
            let aVal, bVal;

            switch(columnIndex) {
                case 0: // State name
                    aVal = a.state;
                    bVal = b.state;
                    break;
                case 1: // Total stores
                    aVal = a.total_stores;
                    bVal = b.total_stores;
                    break;
                case 2: // CATALYST15ORIG
                    aVal = a.sku_counts.CATALYST15ORIG;
                    bVal = b.sku_counts.CATALYST15ORIG;
                    break;
                case 3: // CATALYST15UNSCEN
                    aVal = a.sku_counts.CATALYST15UNSCEN;
                    bVal = b.sku_counts.CATALYST15UNSCEN;
                    break;
                case 4: // CATALYST34LBORIGINAL
                    aVal = a.sku_counts.CATALYST34LBORIGINAL;
                    bVal = b.sku_counts.CATALYST34LBORIGINAL;
                    break;
                case 5: // CATALYSTPET34LBUNSCE
                    aVal = a.sku_counts.CATALYSTPET34LBUNSCE;
                    bVal = b.sku_counts.CATALYSTPET34LBUNSCE;
                    break;
            }

            if (typeof aVal === 'string') {
                return currentSort.ascending ?
                    aVal.localeCompare(bVal) : bVal.localeCompare(aVal);
            } else {
                return currentSort.ascending ? aVal - bVal : bVal - aVal;
            }
        });

        currentData = data;
        populateTable(data);
    }

    // Filter functionality
    document.getElementById('stateFilter').addEventListener('input', (e) => {
        const filter = e.target.value.toLowerCase();
        const filtered = STATE_SUMMARY.states.filter(s =>
            s.state.toLowerCase().includes(filter)
        );
        currentData = filtered;
        populateTable(filtered);
    });

    // Initialize
    populateTable(currentData);
    sortTable(1); // Sort by total stores descending
    </script>
</body>
</html>'''

    return html


def main():
    """Main execution"""
    print("=" * 60)
    print("CATALYST PET STORE MAP - HTML GENERATION")
    print("=" * 60)
    print()

    # Get paths
    project_root = get_project_root()
    input_path = project_root / "data" / "processed" / "stores_geocoded.json"
    readme_output = project_root / "output" / "README.md"

    # Verify input file exists
    if not input_path.exists():
        print(f"ERROR: Input file not found at {input_path}")
        print("Please run 02_geocoding.py first")
        sys.exit(1)

    # Load geocoded stores
    print(f"Loading geocoded stores...")
    stores = load_json(str(input_path))
    geocoded_stores = [s for s in stores if s.get('latitude') and s.get('longitude')]
    print(f"[OK] Loaded {len(stores)} stores ({len(geocoded_stores)} with coordinates)")

    # Create GeoJSON
    print(f"\nCreating GeoJSON...")
    geojson = create_geojson(stores)
    print(f"[OK] Created GeoJSON with {len(geojson['features'])} features")

    # Create state summary
    print(f"\nCreating state summary...")
    state_summary = create_state_summary(stores)
    print(f"[OK] Created summary for {len(state_summary['states'])} states")

    # Generate HTML files
    print(f"\nGenerating HTML files...")

    # Generate combined single-file version
    combined_output = project_root / "output" / "catalyst-store-map.html"
    print(f"  Creating {combined_output.name} (combined map + table)...")
    combined_html = generate_combined_html(geojson, state_summary)
    with open(combined_output, 'w', encoding='utf-8') as f:
        f.write(combined_html)
    print(f"  [OK] Saved: {combined_output}")

    # Generate simplified version (all stores same color)
    simple_output = project_root / "output" / "catalyst-store-map-simple.html"
    print(f"  Creating {simple_output.name} (simple - all stores same color)...")
    simple_html = generate_simple_map_html(geojson, state_summary)
    with open(simple_output, 'w', encoding='utf-8') as f:
        f.write(simple_html)
    print(f"  [OK] Saved: {simple_output}")

    # Generate 3-tab version for web hosting
    web_output = project_root / "output" / "index.html"
    print(f"  Creating {web_output.name} (3-tab web version)...")
    web_html = generate_three_tab_html(geojson, state_summary)
    with open(web_output, 'w', encoding='utf-8') as f:
        f.write(web_html)
    print(f"  [OK] Saved: {web_output}")

    # Create README
    print(f"  Creating {readme_output.name}...")
    readme_content = f"""# Catalyst Pet Store Distribution Map

Interactive map visualization of {geojson['metadata']['total_stores']:,} Walmart stores carrying Catalyst Pet products across the United States.

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

- **Total Stores**: {geojson['metadata']['geocoded_stores']:,}
- **States Covered**: {len(state_summary['states'])}
- **SKU Products**: 4
  - CATALYST15ORIG
  - CATALYST15UNSCEN
  - CATALYST34LBORIGINAL
  - CATALYSTPET34LBUNSCE
- **Unique Combinations**: 11
- **Generated**: {geojson['metadata']['generated_date']}

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
"""

    with open(readme_output, 'w', encoding='utf-8') as f:
        f.write(readme_content)
    print(f"  [OK] Saved: {readme_output}")

    # Print summary
    print(f"\n{'=' * 60}")
    print("[SUCCESS] HTML GENERATION COMPLETE")
    print(f"{'=' * 60}")
    print(f"  Output directory:  {project_root / 'output'}")
    print(f"  Files created:")
    print(f"    - index.html                     (3-TAB WEB VERSION - upload this!)")
    print(f"    - catalyst-store-map.html        (2-view: SKU colors + table)")
    print(f"    - catalyst-store-map-simple.html (2-view: simple + table)")
    print(f"    - README.md                      (documentation)")
    print(f"  ")
    print(f"  Total stores:      {geojson['metadata']['total_stores']:,}")
    print(f"  Geocoded:          {geojson['metadata']['geocoded_stores']:,}")
    print(f"  States:            {len(state_summary['states'])}")
    print(f"  ")
    print(f"  >> FOR WEB HOSTING: Upload {web_output}")
    print(f"  >> Local testing:   Open {combined_output} or {simple_output}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
