"""
Step 1: Data Preparation
Extracts and transforms Excel data into structured JSON format
"""
import pandas as pd
import sys
from pathlib import Path

# Add scripts directory to path for imports
sys.path.append(str(Path(__file__).parent))

from utils import save_json, get_project_root, format_sku_combination, get_sku_color


def prepare_store_data(excel_path):
    """
    Read Excel file and prepare store data

    Args:
        excel_path: Path to Catalyst Store List.xlsx

    Returns:
        List of store dictionaries with SKU combinations
    """
    print("Reading Excel file...")
    df = pd.read_excel(excel_path)

    print(f"[OK] Loaded {len(df)} rows")
    print(f"  Columns: {', '.join(df.columns[:10])}...")

    # Extract relevant columns
    relevant_cols = ['store_number', 'Store Address', 'store_name', 'city_name',
                     'state_or_province_name', 'item_name']
    df = df[relevant_cols].copy()

    print(f"\nProcessing {df['store_number'].nunique()} unique stores...")

    # Group by store to aggregate SKUs
    stores = df.groupby('store_number').agg({
        'Store Address': 'first',
        'store_name': 'first',
        'city_name': 'first',
        'state_or_province_name': 'first',
        'item_name': lambda x: sorted(x.unique())
    }).reset_index()

    # Create SKU combination strings and additional fields
    store_records = []
    for _, row in stores.iterrows():
        skus = row['item_name']
        sku_combination = format_sku_combination(skus)

        store_record = {
            'store_number': int(row['store_number']),
            'store_name': row['store_name'],
            'street_address': row['Store Address'],
            'city': row['city_name'],
            'state': row['state_or_province_name'],
            'skus': skus,
            'sku_combination': sku_combination,
            'sku_count': len(skus),
            'color': get_sku_color(sku_combination),
            'geocode_address': f"{row['Store Address']}, {row['city_name']}, {row['state_or_province_name']}"
        }
        store_records.append(store_record)

    print(f"[OK] Processed {len(store_records)} stores")

    # Print SKU combination summary
    sku_combos = pd.DataFrame(store_records).groupby('sku_combination').size().sort_values(ascending=False)
    print(f"\nSKU Combination Summary:")
    for combo, count in sku_combos.items():
        print(f"  {count:4d} stores: {combo}")

    return store_records


def main():
    """Main execution"""
    print("=" * 60)
    print("CATALYST PET STORE MAP - DATA PREPARATION")
    print("=" * 60)
    print()

    # Get paths
    project_root = get_project_root()
    excel_path = project_root / "Catalyst Store List.xlsx"
    output_path = project_root / "data" / "processed" / "stores_prepared.json"

    # Verify Excel file exists
    if not excel_path.exists():
        print(f"ERROR: Excel file not found at {excel_path}")
        sys.exit(1)

    # Process data
    stores = prepare_store_data(excel_path)

    # Save to JSON
    print(f"\nSaving to {output_path}...")
    save_json(stores, str(output_path))

    print(f"\n{'=' * 60}")
    print("[SUCCESS] DATA PREPARATION COMPLETE")
    print(f"  Output: {output_path}")
    print(f"  Stores: {len(stores)}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
