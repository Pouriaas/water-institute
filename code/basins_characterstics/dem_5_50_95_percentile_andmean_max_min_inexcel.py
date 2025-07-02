#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Jun 20 20:42:50 2025

@author: pouria
"""

import os
import glob
import numpy as np
import rasterio
import pandas as pd

# Folder with clipped NDVI TIFFs
ndvi_folder = '/home/pouria/git/water-institute/data/basins_charactristics/output/dem_cliped415'
output_excel = '/home/pouria/git/water-institute/data/basins_charactristics/output/excel/41/415/dem_415.xlsx'
# os.makedirs(output_excel, exist_ok=True)
# Collect results
results = []

# Loop through each TIFF
for tif_path in glob.glob(os.path.join(ndvi_folder, 'ndvi_*.tif')):
    point_id = os.path.basename(tif_path).split('_')[1].split('.')[0]

    with rasterio.open(tif_path) as src:
        data = src.read(1)
        data = data[data != src.nodata]  # Remove NoData values
        if data.size == 0:
            print(f'⚠️ No valid data in {point_id}, skipping...')
            continue

        # Compute statistics
        # mean = np.mean(data)
        # mini = np.min(data)
        # maxi = np.max(data)
        p5 = np.percentile(data, 5)
        p50 = np.percentile(data, 50)
        p95 = np.percentile(data, 95)

        results.append({
            'Point_ID': int(point_id),
            # 'min': round(mini, 4),
            # 'Mean': round(mean, 4),
            # 'max': round(maxi, 4),
            'P5': round(p5, 4),
            'P50': round(p50, 4),
            'P95': round(p95, 4)
        })

# Save to Excel
df = pd.DataFrame(results)
df.sort_values(by='Point_ID', inplace=True)
df.to_excel(output_excel, index=False)

print(f'✅ NDVI stats saved to {output_excel}')
