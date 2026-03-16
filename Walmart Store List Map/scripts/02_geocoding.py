"""
Step 2: Geocoding
Converts store street addresses to latitude/longitude using the US Census Geocoder batch API.
Much faster than Nominatim — processes all addresses in a single batch request.
"""
import requests
import io
import csv
import time
import sys
from pathlib import Path
from datetime import datetime

# Add scripts directory to path for imports
sys.path.append(str(Path(__file__).parent))

from utils import load_json, save_json, get_project_root


def build_batch_csv(stores):
    """
    Build CSV string for the Census Geocoder batch API.

    Format: Unique ID, Street Address, City, State, ZIP (blank)
    """
    output = io.StringIO()
    writer = csv.writer(output)
    for store in stores:
        writer.writerow([
            store['store_number'],
            store.get('street_address', ''),
            store.get('city', ''),
            store.get('state', ''),
            ''  # ZIP — we don't have it, but the field is required
        ])
    return output.getvalue()


def parse_batch_response(response_text):
    """
    Parse the Census Geocoder batch response CSV.

    Returns dict mapping store_number -> {'lat': float, 'lon': float} or None
    """
    results = {}
    reader = csv.reader(io.StringIO(response_text))
    for row in reader:
        if len(row) < 6:
            continue
        store_number = int(row[0].strip('"'))
        match_type = row[2].strip('"').strip()

        if match_type in ('Match', 'Exact'):
            # Coordinates are in the format "lon,lat" in the coordinate field
            coords_str = row[5].strip('"').strip()
            if coords_str:
                try:
                    lon, lat = coords_str.split(',')
                    results[store_number] = {
                        'lat': float(lat),
                        'lon': float(lon)
                    }
                except (ValueError, IndexError):
                    results[store_number] = None
            else:
                results[store_number] = None
        else:
            results[store_number] = None

    return results


def geocode_batch(stores, benchmark='Public_AR_Current', chunk_size=1000):
    """
    Geocode stores using the US Census Geocoder batch API.

    The API accepts up to 10,000 addresses per request. We chunk into
    groups of 1,000 for reliability and progress reporting.

    Args:
        stores: List of store dictionaries
        benchmark: Census geocoder benchmark to use
        chunk_size: Number of addresses per batch request

    Returns:
        Dict mapping store_number -> {'lat', 'lon'} or None
    """
    url = 'https://geocoding.geo.census.gov/geocoder/locations/addressbatch'
    all_results = {}

    chunks = [stores[i:i + chunk_size] for i in range(0, len(stores), chunk_size)]
    print(f"  Split into {len(chunks)} batch(es) of up to {chunk_size} addresses")

    for i, chunk in enumerate(chunks, 1):
        print(f"\n  Batch {i}/{len(chunks)} ({len(chunk)} addresses)...")
        batch_csv = build_batch_csv(chunk)

        files = {
            'addressFile': ('addresses.csv', batch_csv, 'text/csv')
        }
        data = {
            'benchmark': benchmark,
        }

        retries = 3
        for attempt in range(retries):
            try:
                response = requests.post(url, files=files, data=data, timeout=300)
                response.raise_for_status()
                results = parse_batch_response(response.text)
                all_results.update(results)

                matched = sum(1 for v in results.values() if v is not None)
                print(f"    Matched: {matched}/{len(chunk)} ({matched/len(chunk)*100:.1f}%)")
                break

            except Exception as e:
                if attempt < retries - 1:
                    wait = 10 * (attempt + 1)
                    print(f"    Error: {e}. Retrying in {wait}s...")
                    time.sleep(wait)
                else:
                    print(f"    ERROR: Batch {i} failed after {retries} attempts: {e}")
                    # Mark all in this chunk as failed
                    for store in chunk:
                        all_results[store['store_number']] = None

    return all_results


def nominatim_fallback(failed_stores):
    """
    Fallback geocoder for addresses the Census API couldn't match.
    Uses Nominatim (OpenStreetMap) with city/state only for a rough location.

    Args:
        failed_stores: List of store dicts that failed Census geocoding

    Returns:
        Dict mapping store_number -> {'lat', 'lon'} or None
    """
    if not failed_stores:
        return {}

    print(f"\n  Falling back to Nominatim for {len(failed_stores)} failed addresses...")
    print(f"  (Using city/state for approximate location, ~{len(failed_stores) * 1.1 / 60:.1f} min)")

    url = "https://nominatim.openstreetmap.org/search"
    headers = {'User-Agent': 'CatalystPetStoreMapper/1.0 (Educational Project)'}
    results = {}

    # Deduplicate by city/state to minimize API calls
    city_state_map = {}
    for store in failed_stores:
        key = (store['city'], store['state'])
        if key not in city_state_map:
            city_state_map[key] = []
        city_state_map[key].append(store['store_number'])

    print(f"  Unique city/state combos to look up: {len(city_state_map)}")

    resolved = 0
    for (city, state), store_numbers in city_state_map.items():
        query = f"{city}, {state}, USA"
        params = {'q': query, 'format': 'json', 'limit': 1, 'countrycodes': 'us'}

        try:
            response = requests.get(url, params=params, headers=headers, timeout=10)
            response.raise_for_status()
            data = response.json()

            if data:
                coords = {'lat': float(data[0]['lat']), 'lon': float(data[0]['lon'])}
                for sn in store_numbers:
                    results[sn] = coords
                resolved += len(store_numbers)
        except Exception:
            for sn in store_numbers:
                results[sn] = None
        finally:
            time.sleep(1.1)

    print(f"  Nominatim recovered: {resolved}/{len(failed_stores)} stores")
    return results


def main():
    """Main execution"""
    print("=" * 60)
    print("CATALYST PET STORE MAP - GEOCODING (US Census Batch API)")
    print("=" * 60)
    print()

    # Get paths
    project_root = get_project_root()
    input_path = project_root / "data" / "processed" / "stores_prepared.json"
    output_path = project_root / "data" / "processed" / "stores_geocoded.json"
    cache_path = project_root / "data" / "processed" / "geocoding_cache.json"
    failures_path = project_root / "data" / "processed" / "geocoding_failures.txt"

    # Verify input file exists
    if not input_path.exists():
        print(f"ERROR: Input file not found at {input_path}")
        print("Please run 01_data_preparation.py first")
        sys.exit(1)

    # Load stores
    print(f"Loading stores from {input_path.name}...")
    stores = load_json(str(input_path))
    print(f"[OK] Loaded {len(stores)} stores")

    # Geocode via Census batch API
    print(f"\nGeocoding {len(stores)} stores via US Census Geocoder...")
    start_time = time.time()
    geocode_results = geocode_batch(stores)
    elapsed_time = time.time() - start_time

    # Assign coordinates to stores — first pass (Census results)
    geocoded_count = 0
    failed_stores_for_fallback = []

    for store in stores:
        coords = geocode_results.get(store['store_number'])
        if coords:
            store['latitude'] = coords['lat']
            store['longitude'] = coords['lon']
            geocoded_count += 1
        else:
            store['latitude'] = None
            store['longitude'] = None
            failed_stores_for_fallback.append(store)

    # Nominatim fallback for Census failures (city/state level)
    if failed_stores_for_fallback:
        fallback_results = nominatim_fallback(failed_stores_for_fallback)
        for store in failed_stores_for_fallback:
            coords = fallback_results.get(store['store_number'])
            if coords:
                store['latitude'] = coords['lat']
                store['longitude'] = coords['lon']
                geocode_results[store['store_number']] = coords
                geocoded_count += 1

    # Collect final failures
    failed_stores = []
    for store in stores:
        if store['latitude'] is None:
            failed_stores.append({
                'store_number': store['store_number'],
                'address': store['geocode_address']
            })

    # Save geocoded stores
    print(f"\nSaving geocoded data to {output_path.name}...")
    save_json(stores, str(output_path))

    # Save cache for future reference (flat dict keyed by store_number)
    cache_data = {str(k): v for k, v in geocode_results.items()}
    save_json(cache_data, str(cache_path), indent=2)

    # Save failures if any
    if failed_stores:
        print(f"\nSaving {len(failed_stores)} failed geocoding attempts...")
        with open(failures_path, 'w', encoding='utf-8') as f:
            f.write(f"Failed Geocoding - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 60 + "\n\n")
            for failure in failed_stores:
                f.write(f"Store #{failure['store_number']}: {failure['address']}\n")
        print(f"[OK] Saved failures to {failures_path.name}")

    # Print statistics
    print(f"\n{'=' * 60}")
    print("[SUCCESS] GEOCODING COMPLETE")
    print(f"{'=' * 60}")
    print(f"  Total stores:          {len(stores)}")
    print(f"  Successfully geocoded: {geocoded_count} ({geocoded_count/len(stores)*100:.1f}%)")
    print(f"  Failed:                {len(failed_stores)} ({len(failed_stores)/len(stores)*100:.1f}%)")
    print(f"  Time elapsed:          {elapsed_time:.1f} seconds")
    print(f"  Output:                {output_path}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
