#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jun  2 13:44:14 2025

@author: pouria
"""
import os
import rasterio
import geopandas as gpd
from rasterio.mask import mask
from shapely.geometry import mapping
from joblib import Parallel, delayed

# === Paths ===
input_tif_dir = "/home/pouria/git/water-institute/data/basins_charactristics/input/soil_grids"
shapefile_path = "/home/pouria/git/water-institute/data/basins_charactristics/input/shapefiles/41 AllPoint SubWatershed/415_SubWatershed_AllPoint_4040226_Merged.shp"
output_dir = "/home/pouria/git/water-institute/data/basins_charactristics/output/soilgrids_cliped/415"
os.makedirs(output_dir, exist_ok=True)
n_jobs = 5 # Use all CPU cores

# === Load Shapefile ===
gdf = gpd.read_file(shapefile_path)


def process_single_clip(tif_path, tif_name, row, gdf_crs):
    Point_ID = str(row["Point_ID"])
    geometry = [mapping(row["geometry"])]

    # Output folder for this tif file
    tif_output_dir = os.path.join(output_dir, tif_name)
    os.makedirs(tif_output_dir, exist_ok=True)

    try:
        with rasterio.open(tif_path) as src:
            # Reproject shapefile geometry to raster CRS if needed
            if gdf_crs != src.crs:
                row = row.to_frame().T.set_crs(gdf_crs).to_crs(src.crs).iloc[0]
                geometry = [mapping(row["geometry"])]

            out_image, out_transform = mask(src, geometry, crop=True)
            out_meta = src.meta.copy()
            out_meta.update({
                "driver": "GTiff",
                "height": out_image.shape[1],
                "width": out_image.shape[2],
                "transform": out_transform
            })

            out_path = os.path.join(tif_output_dir, f"{Point_ID}.tif")
            with rasterio.open(out_path, "w", **out_meta) as dest:
                dest.write(out_image)

    except Exception as e:
        print(f"❌ Skipped {Point_ID} for {tif_name}.tif: {e}")

def process_single_tif(tif_file):
    if not tif_file.endswith(".tif"):
        return
    tif_path = os.path.join(input_tif_dir, tif_file)
    tif_name = os.path.splitext(tif_file)[0]
    print(f"🔄 Processing {tif_file}...")

    Parallel(n_jobs=n_jobs)(
        delayed(process_single_clip)(tif_path, tif_name, row, gdf.crs)
        for _, row in gdf.iterrows()
    )

    print(f"✅ Finished {tif_file}")

# === Run for all .tif files ===
all_tifs = [f for f in os.listdir(input_tif_dir) if f.endswith(".tif")]
Parallel(n_jobs=n_jobs)(delayed(process_single_tif)(tif_file) for tif_file in all_tifs)

print("🏁 All clipping finished.")