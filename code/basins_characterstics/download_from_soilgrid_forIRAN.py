#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Jun  1 09:33:02 2025
Final version: Download, reproject, and clip SoilGrids to Iran (WGS84)
"""

import os
from osgeo import gdal
from tqdm import tqdm

# Enable GDAL exceptions
gdal.UseExceptions()

# --- Configuration ---

# Output directory
output_dir = "/home/pouria/git/water-institute/data/basins_charactristics/input/soil_grids"
os.makedirs(output_dir, exist_ok=True)

# Iran's bounding box in WGS84 (lon/lat)
iran_bbox = [44.0, 25.0, 64.0, 40.0]  # min_lon, min_lat, max_lon, max_lat

# Soil properties to download
properties = ["clay", "silt", "sand"]

# Depths in cm
depths = {
    0: "0-5cm",
    5: "5-15cm",
    15: "15-30cm",
    30: "30-60cm",
    60: "60-100cm",
    100: "100-200cm"
}

# Quantiles/statistics to download
quantiles = ["mean", "Q0.05", "Q0.50", "Q0.95"]

# Base URL for SoilGrids data
base_url = "https://files.isric.org/soilgrids/latest/data/"

# --- Download Loop ---
print(f"📥 Starting download of SoilGrids data for Iran into '{output_dir}'...\n")

total = len(properties) * len(depths) * len(quantiles)
progress = tqdm(total=total)

for prop in properties:
    for depth_key, depth_str in depths.items():
        for q in quantiles:
            filename = f"{prop}_{depth_str}_{q}.tif"
            output_path = os.path.join(output_dir, filename)

            # Build the remote VRT URL via /vsicurl
            vrt_url = f"/vsicurl?url={base_url}{prop}/{prop}_{depth_str}_{q}.vrt"

            try:
                # Reproject and clip to Iran in WGS84
                gdal.Warp(
                    destNameOrDestDS=output_path,
                    srcDSOrSrcDSTab=vrt_url,
                    format="GTiff",
                    dstSRS="EPSG:4326",
                    outputBounds=iran_bbox,
                    creationOptions=["TILED=YES", "COMPRESS=DEFLATE", "PREDICTOR=2", "BIGTIFF=YES"],
                    multithread=True
                )
            except Exception as e:
                print(f"❌ Failed: {filename} | Reason: {e}")

            progress.update(1)

progress.close()
print("\n✅ Download, reproject, and clipping complete!")
print(f"✔️ Files saved in: {output_dir}")