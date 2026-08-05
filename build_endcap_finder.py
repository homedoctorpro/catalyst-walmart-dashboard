"""
Public Endcap Store Finder — consumer-facing page for external sharing.

Reads endcap_stores.csv (produced by endcap_export.py) and writes
endcap_finder.html: a self-contained Leaflet map of all Catalyst Pet
Walmart endcap stores with a zip-code search that lists the nearest
stores (distance + Google Maps directions link).

Unlike endcap_map.html (internal), this page carries no internal fields —
no "no Catalyst yet" flags, no endcap quantities, no SKU trait data.

Refresh after endcap_export.py:
  python build_endcap_finder.py

Deploy: commit endcap_finder.html and push to master (GitHub Pages).
Public URL:
  https://homedoctorpro.github.io/catalyst-walmart-dashboard/endcap_finder.html
"""
from __future__ import annotations

import csv
import json
from pathlib import Path

ROOT = Path(__file__).parent
IN_CSV = ROOT / "endcap_stores.csv"
OUT_HTML = ROOT / "endcap_finder.html"


def load_stores() -> list[dict]:
    stores = []
    with open(IN_CSV, encoding="utf-8") as f:
        for r in csv.DictReader(f):
            if not r["latitude"] or not r["longitude"]:
                continue
            skus = [k for k, col in (("15O", "has_15O"), ("34O", "has_34O"),
                                     ("15U", "has_15U"), ("34U", "has_34U")) if r[col] == "Y"]
            stores.append({
                "n": int(r["store_number"]),
                "a": r["street_address"],
                "c": r["city"].title() if r["city"] else "",
                "s": r["state"],
                "z": r["zip"],
                "k": skus,
                "lat": round(float(r["latitude"]), 5),
                "lon": round(float(r["longitude"]), 5),
            })
    return stores


def main():
    stores = load_stores()
    html = TEMPLATE.replace("/*STORES_PLACEHOLDER*/", json.dumps(stores, separators=(",", ":")))
    OUT_HTML.write_text(html, encoding="utf-8")
    print(f"Wrote {OUT_HTML.name}: {len(stores)} stores embedded ({OUT_HTML.stat().st_size:,} bytes)")


TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Catalyst Pet — Walmart Endcap Store Finder</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="description" content="Find the Walmart store nearest you featuring Catalyst Pet on an endcap display. Enter your zip code to see the closest locations.">
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css">
<link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css">
<link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css">
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<script src="https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js"></script>
<style>
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; height: 100%; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; color: #222; background: #f7f8fa; }
  #header { background: #1a3c34; color: #fff; padding: 16px 22px; }
  #header h1 { margin: 0; font-size: 1.2rem; font-weight: 700; }
  #header .sub { font-size: 0.85rem; opacity: 0.85; margin-top: 4px; }

  #searchbar { background: #fff; padding: 12px 22px; border-bottom: 1px solid #e0e0e0; display: flex; gap: 10px; flex-wrap: wrap; align-items: center; }
  #searchbar input[type=text] { padding: 10px 14px; border: 1px solid #ccc; border-radius: 8px; font-size: 1rem; width: 180px; }
  #searchbar button { padding: 10px 18px; border: none; border-radius: 8px; font-size: 0.95rem; font-weight: 700; cursor: pointer; }
  #btn-find { background: #1a3c34; color: #fff; }
  #btn-find:hover { background: #275a4e; }
  #btn-geo { background: #eef2f1; color: #1a3c34; }
  #btn-geo:hover { background: #e0e8e6; }
  #search-msg { font-size: 0.85rem; color: #b00020; flex-basis: 100%; }
  #search-msg.ok { color: #1a3c34; }

  #content { display: flex; height: calc(100vh - 148px); min-height: 420px; }
  #map { flex: 1; height: 100%; }
  #results { width: 360px; max-width: 45%; overflow-y: auto; background: #fff; border-left: 1px solid #e0e0e0; display: none; }
  #results.visible { display: block; }
  #results h2 { font-size: 0.95rem; margin: 0; padding: 14px 16px 10px; color: #1a3c34; border-bottom: 1px solid #eee; position: sticky; top: 0; background: #fff; }
  .store-card { padding: 12px 16px; border-bottom: 1px solid #f0f0f0; cursor: pointer; }
  .store-card:hover { background: #f4f8f7; }
  .store-card .rank { display: inline-flex; align-items: center; justify-content: center; width: 22px; height: 22px; border-radius: 50%; background: #1a3c34; color: #fff; font-size: 0.75rem; font-weight: 700; margin-right: 8px; }
  .store-card .dist { float: right; font-weight: 700; color: #1a3c34; font-size: 0.9rem; }
  .store-card .name { font-weight: 700; font-size: 0.92rem; }
  .store-card .addr { font-size: 0.85rem; color: #555; margin: 4px 0 6px 30px; }
  .store-card a.dir { margin-left: 30px; font-size: 0.82rem; font-weight: 700; color: #0b57d0; text-decoration: none; }
  .store-card a.dir:hover { text-decoration: underline; }

  .leaflet-popup-content { font-size: 0.85rem; line-height: 1.5; min-width: 190px; }
  .leaflet-popup-content b { color: #1a3c34; }
  .leaflet-popup-content a { color: #0b57d0; font-weight: 700; text-decoration: none; }

  .sku-chip { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 0.72rem; font-weight: 700; color: #fff; margin: 2px 3px 0 0; white-space: nowrap; }
  .sku-chip.s-15O { background: #1f6f43; }
  .sku-chip.s-34O { background: #2e7d5b; }
  .sku-chip.s-15U { background: #4d8a9e; }
  .sku-chip.s-34U { background: #6a5c9e; }
  .sku-none { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 0.72rem; font-weight: 700; color: #666; background: #eee; margin-top: 2px; }
  .sku-line { margin-top: 4px; }
  .sku-label { font-size: 0.75rem; color: #888; font-weight: 700; text-transform: uppercase; letter-spacing: 0.03em; }

  #footer { background: #fff; border-top: 1px solid #e0e0e0; padding: 8px 22px; font-size: 0.75rem; color: #888; }

  @media (max-width: 700px) {
    #content { flex-direction: column; height: auto; }
    #map { height: 48vh; }
    #results { width: 100%; max-width: 100%; border-left: none; border-top: 1px solid #e0e0e0; max-height: 45vh; }
    #searchbar input[type=text] { flex: 1; width: auto; }
  }
</style>
</head>
<body>
<div id="header">
  <h1>Catalyst Pet — Walmart Endcap Store Finder</h1>
  <div class="sub" id="sub">Enter your zip code to find the nearest Walmart featuring Catalyst Pet on an endcap display.</div>
</div>

<div id="searchbar">
  <input type="text" id="zip" inputmode="numeric" maxlength="5" placeholder="Zip code" autocomplete="postal-code">
  <button id="btn-find">Find stores</button>
  <button id="btn-geo">📍 Use my location</button>
  <span id="search-msg"></span>
</div>

<div id="content">
  <div id="map"></div>
  <div id="results">
    <h2 id="results-title"></h2>
    <div id="results-list"></div>
  </div>
</div>

<div id="footer">
  <span id="store-count"></span> Walmart locations featuring Catalyst Pet endcap displays. Product availability may vary by store.
</div>

<script>
const STORES = /*STORES_PLACEHOLDER*/;
const NEAREST_N = 10;

document.getElementById("store-count").textContent = STORES.length.toLocaleString();

// ─── Map ───────────────────────────────────────────────────────────────────
const map = L.map("map", { preferCanvas: true }).setView([39.5, -98.35], 4);
L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
  attribution: "© OpenStreetMap", maxZoom: 19,
}).addTo(map);

function directionsUrl(st) {
  const dest = encodeURIComponent(`Walmart, ${st.a}, ${st.c}, ${st.s} ${st.z}`);
  return `https://www.google.com/maps/dir/?api=1&destination=${dest}`;
}
const SKU_NAMES = { "15O": "15 lb Original", "34O": "34 lb Original", "15U": "15 lb Unscented", "34U": "34 lb Unscented" };
function skuChips(st) {
  return st.k.length
    ? st.k.map(k => `<span class="sku-chip s-${k}">${SKU_NAMES[k]}</span>`).join("")
    : `<span class="sku-none">None</span>`;
}
function popupHtml(st, distMi) {
  const addr = [st.a, st.c, st.s, st.z].filter(Boolean).join(", ");
  return `<b>Walmart #${st.n}</b><br>${addr}` +
    (distMi != null ? `<br><span style="color:#1a3c34;font-weight:700">${distMi.toFixed(1)} miles away</span>` : "") +
    `<div class="sku-line"><span class="sku-label">Originally stocked:</span><br>${skuChips(st)}</div>` +
    `<a href="${directionsUrl(st)}" target="_blank" rel="noopener">Get directions →</a>`;
}

const cluster = L.markerClusterGroup({
  chunkedLoading: true, showCoverageOnHover: false, maxClusterRadius: 45,
  iconCreateFunction: c => L.divIcon({
    html: `<div style="background:#1a3c34;color:#fff;border-radius:50%;width:32px;height:32px;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:0.75rem;border:2px solid #fff;box-shadow:0 1px 3px rgba(0,0,0,.3)">${c.getChildCount()}</div>`,
    className: "", iconSize: [32, 32],
  }),
});
const markerByStore = {};
STORES.forEach(st => {
  const m = L.circleMarker([st.lat, st.lon], {
    radius: 6, color: "#1a3c34", fillColor: "#2e7d5b", fillOpacity: 0.85, weight: 1.5,
  });
  m.bindPopup(popupHtml(st, null));
  markerByStore[st.n] = m;
  cluster.addLayer(m);
});
map.addLayer(cluster);

// ─── Zip lookup ────────────────────────────────────────────────────────────
// First try the zips we already know (store zips), then fall back to the free
// zippopotam.us API for any other US zip.
const zipIndex = {};
STORES.forEach(st => { if (st.z && !(st.z in zipIndex)) zipIndex[st.z] = [st.lat, st.lon]; });

async function geocodeZip(zip) {
  if (zipIndex[zip]) return zipIndex[zip];
  const resp = await fetch(`https://api.zippopotam.us/us/${zip}`);
  if (!resp.ok) throw new Error("zip not found");
  const data = await resp.json();
  const p = data.places && data.places[0];
  if (!p) throw new Error("zip not found");
  return [parseFloat(p.latitude), parseFloat(p.longitude)];
}

// ─── Distance + results ────────────────────────────────────────────────────
function haversineMiles(lat1, lon1, lat2, lon2) {
  const R = 3958.8, toRad = d => d * Math.PI / 180;
  const dLat = toRad(lat2 - lat1), dLon = toRad(lon2 - lon1);
  const a = Math.sin(dLat / 2) ** 2 +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(a));
}

let originMarker = null;
let rankMarkers = [];

function showNearest(lat, lon, label) {
  const ranked = STORES
    .map(st => ({ st, d: haversineMiles(lat, lon, st.lat, st.lon) }))
    .sort((a, b) => a.d - b.d)
    .slice(0, NEAREST_N);

  if (originMarker) map.removeLayer(originMarker);
  rankMarkers.forEach(m => map.removeLayer(m));
  rankMarkers = [];

  originMarker = L.marker([lat, lon], {
    icon: L.divIcon({
      html: `<div style="background:#d32f2f;width:18px;height:18px;border-radius:50%;border:3px solid #fff;box-shadow:0 1px 4px rgba(0,0,0,.4)"></div>`,
      className: "", iconSize: [18, 18], iconAnchor: [9, 9],
    }),
  }).addTo(map).bindPopup(`<b>${label}</b>`);

  ranked.forEach((r, i) => {
    const m = L.marker([r.st.lat, r.st.lon], {
      icon: L.divIcon({
        html: `<div style="background:#1a3c34;color:#fff;width:26px;height:26px;border-radius:50%;border:2px solid #fff;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:0.8rem;box-shadow:0 1px 4px rgba(0,0,0,.4)">${i + 1}</div>`,
        className: "", iconSize: [26, 26], iconAnchor: [13, 13],
      }),
      zIndexOffset: 1000,
    }).addTo(map).bindPopup(popupHtml(r.st, r.d));
    rankMarkers.push(m);
  });

  const bounds = L.latLngBounds([[lat, lon]].concat(ranked.slice(0, 5).map(r => [r.st.lat, r.st.lon])));
  map.fitBounds(bounds.pad(0.25));

  document.getElementById("results").classList.add("visible");
  document.getElementById("results-title").textContent = `Nearest stores to ${label}`;
  document.getElementById("results-list").innerHTML = ranked.map((r, i) => `
    <div class="store-card" data-store="${r.st.n}">
      <span class="dist">${r.d.toFixed(1)} mi</span>
      <span class="rank">${i + 1}</span><span class="name">Walmart #${r.st.n}</span>
      <div class="addr">${[r.st.a, r.st.c, r.st.s, r.st.z].filter(Boolean).join(", ")}</div>
      <div class="addr sku-line"><span class="sku-label">Originally stocked:</span> ${skuChips(r.st)}</div>
      <a class="dir" href="${directionsUrl(r.st)}" target="_blank" rel="noopener">Get directions →</a>
    </div>
  `).join("");

  document.querySelectorAll(".store-card").forEach((card, i) => {
    card.addEventListener("click", e => {
      if (e.target.tagName === "A") return;
      const m = rankMarkers[i];
      map.setView(m.getLatLng(), 12);
      m.openPopup();
    });
  });
  map.invalidateSize();
}

// ─── Search handlers ───────────────────────────────────────────────────────
const msg = document.getElementById("search-msg");

async function findByZip() {
  const zip = document.getElementById("zip").value.trim();
  msg.textContent = ""; msg.className = "";
  if (!/^\d{5}$/.test(zip)) {
    msg.textContent = "Please enter a 5-digit zip code.";
    return;
  }
  try {
    const [lat, lon] = await geocodeZip(zip);
    showNearest(lat, lon, `zip ${zip}`);
  } catch (e) {
    msg.textContent = "Sorry, we couldn't find that zip code. Please check it and try again.";
  }
}

document.getElementById("btn-find").addEventListener("click", findByZip);
document.getElementById("zip").addEventListener("keydown", e => { if (e.key === "Enter") findByZip(); });

document.getElementById("btn-geo").addEventListener("click", () => {
  msg.textContent = ""; msg.className = "";
  if (!navigator.geolocation) {
    msg.textContent = "Location is not available in this browser — please enter your zip code.";
    return;
  }
  msg.textContent = "Locating…"; msg.className = "ok";
  navigator.geolocation.getCurrentPosition(
    pos => {
      msg.textContent = "";
      showNearest(pos.coords.latitude, pos.coords.longitude, "your location");
    },
    () => {
      msg.textContent = "We couldn't get your location — please enter your zip code instead.";
      msg.className = "";
    },
    { timeout: 10000 }
  );
});
</script>
</body>
</html>
"""


if __name__ == "__main__":
    main()
