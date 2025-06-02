#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jun  2 10:16:26 2025

@author: pouria
"""

import rasterio
import geopandas as gpd
import numpy as np
import pandas as pd
from rasterio.features import geometry_mask
from rasterio.mask import mask
from collections import Counter
import os

# Input paths
tif_folder = "path/to/tif_files"
basin_shapefile = "path/to/basin_shapefile.shp"  # Must have 'Point_Id' field

# Output Excel
output_excel = "basin_class_summary.xlsx"

# Read basins shapefile
gdf = gpd.read_file(basin_shapefile)

# Initialize output dataframe
all_results = []

# Loop through each .tif file
for tif_file in os.listdir(tif_folder):
    if not tif_file.endswith('.tif'):
        continue

    with rasterio.open(os.path.join(tif_folder, tif_file)) as src:
        for idx, basin in gdf.iterrows():
            geom = [basin['geometry']]
            point_id = basin['Point_Id']

            try:
                out_image, out_transform = mask(src, geom, crop=True)
            except Exception as e:
                print(f"Failed to mask for basin {point_id} in {tif_file}: {e}")
                continue

            data = out_image[0]
            data = data[data > 0]  # Remove no-data (assumed to be 0)

            # Count pixel classes
            class_counts = dict(Counter(data.flatten()))
            total_pixels = sum(class_counts.values())

            # Normalize to percentage
            class_percent = {f"class{int(k)}": round((v / total_pixels) * 100, 2) 
                             for k, v in class_counts.items()}

            class_percent["Point_Id"] = point_id
            class_percent["tif_name"] = tif_file

            all_results.append(class_percent)

# Convert to DataFrame
df = pd.DataFrame(all_results)

# Fill missing class columns with 0
for i in range(1, 14):
    col = f'class{i}'
    if col not in df.columns:
        df[col] = 0
df.fillna(0, inplace=True)

# Reorder columns
cols = ['Point_Id', 'tif_name'] + [f'class{i}' for i in range(1, 14)]
df = df[cols]

# Save to Excel
df.to_excel(output_excel, index=False)
print(f"Saved result to {output_excel}")