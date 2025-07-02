
import rasterio
import numpy as np
import pandas as pd
from collections import Counter
import os
import re

# Class value to label mapping
class_labels = {
    1: "Urban",
    2: "Water",
    3: "Wetland",
    4: "Kalut_yardang",
    5: "Marshland",
    6: "Salty_Land",
    7: "Clay",
    8: "Forest",
    9: "Outcrop",
    10: "Uncovered_Plain",
    11: "Sand",
    12: "Farm_Land",
    13: "Range_Land"
}

tif_folder = "/home/pouria/git/water-institute/data/basins_charactristics/output/land_cover/415"
output_excel = "/home/pouria/git/water-institute/data/basins_charactristics/output/excel/41/415/land_cover.xlsx"

# Initialize results list
results = []

# Loop through TIFs
for filename in os.listdir(tif_folder):
    if not filename.endswith('.tif'):
        continue

    # Extract Point_Id from filename
    match = re.search(r'(\d+)', filename)
    if not match:
        print(f"Could not extract Point_Id from {filename}")
        continue
    point_id = int(match.group(1))

    # Open TIF and read data
    tif_path = os.path.join(tif_folder, filename)
    with rasterio.open(tif_path) as src:
        data = src.read(1)

    # Filter out no-data (assumed to be 0)
    valid_pixels = data[data > 0]

    # Count pixel values
    pixel_counts = dict(Counter(valid_pixels.flatten()))
    total_pixels = sum(pixel_counts.values())

    # Compute percentages using class labels
    class_percentages = {
        class_labels[i]: round((pixel_counts.get(i, 0) / total_pixels) * 100, 2)
        for i in range(1, 14)
    }

    # Add Point_Id to result
    class_percentages["Point_Id"] = point_id
    results.append(class_percentages)

# Create DataFrame
df = pd.DataFrame(results)

# Ensure all class columns are present
for label in class_labels.values():
    if label not in df.columns:
        df[label] = 0.0

# Reorder columns
ordered_cols = ["Point_Id"] + list(class_labels.values())
df = df[ordered_cols]

# Save to Excel
df.to_excel(output_excel, index=False)
print(f"Saved results to {output_excel}")