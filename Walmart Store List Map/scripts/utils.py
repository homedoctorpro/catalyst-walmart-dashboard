"""
Shared utilities for Catalyst Pet store mapping project
"""
import json
import os
from pathlib import Path

# SKU color mapping for 11 unique combinations
SKU_COLORS = {
    # Single SKU (3 colors - lighter shades)
    "CATALYST15ORIG": "#4CAF50",                                    # Green (most common)
    "CATALYST34LBORIGINAL": "#2196F3",                             # Blue
    "CATALYSTPET34LBUNSCE": "#FF9800",                             # Orange

    # Two SKUs (5 colors - medium intensity)
    "CATALYST15ORIG, CATALYST34LBORIGINAL": "#009688",             # Teal
    "CATALYST15ORIG, CATALYST15UNSCEN": "#8BC34A",                 # Light Green
    "CATALYST15ORIG, CATALYSTPET34LBUNSCE": "#CDDC39",             # Lime
    "CATALYST15UNSCEN, CATALYST34LBORIGINAL": "#3F51B5",           # Indigo
    "CATALYST15UNSCEN, CATALYSTPET34LBUNSCE": "#FF5722",           # Deep Orange

    # Three SKUs (2 colors - darker shades)
    "CATALYST15ORIG, CATALYST15UNSCEN, CATALYST34LBORIGINAL": "#673AB7",  # Deep Purple
    "CATALYST15ORIG, CATALYST15UNSCEN, CATALYSTPET34LBUNSCE": "#E91E63",  # Pink

    # All 4 SKUs (1 color - darkest/special)
    "CATALYST15ORIG, CATALYST15UNSCEN, CATALYST34LBORIGINAL, CATALYSTPET34LBUNSCE": "#F44336"  # Red (premium)
}


def get_sku_color(sku_combination):
    """Get color for SKU combination"""
    return SKU_COLORS.get(sku_combination, "#999999")


def save_json(data, filepath, indent=2):
    """Save data to JSON file with pretty formatting"""
    # Ensure directory exists
    os.makedirs(os.path.dirname(filepath), exist_ok=True)

    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=indent, ensure_ascii=False)
    print(f"[OK] Saved: {filepath}")


def load_json(filepath):
    """Load data from JSON file"""
    with open(filepath, 'r', encoding='utf-8') as f:
        return json.load(f)


def format_sku_combination(skus):
    """Format list of SKUs into sorted comma-separated string"""
    return ', '.join(sorted(skus))


def get_project_root():
    """Get the project root directory"""
    return Path(__file__).parent.parent


def get_data_path(filename):
    """Get full path for data file"""
    return get_project_root() / 'data' / 'processed' / filename


def get_output_path(filename):
    """Get full path for output file"""
    return get_project_root() / 'output' / filename
