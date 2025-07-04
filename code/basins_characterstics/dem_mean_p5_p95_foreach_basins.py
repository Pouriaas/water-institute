#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue May 27 12:30:01 2025

@author: pouria
"""

import geopandas as gpd
import os
import pandas as pd
import warnings; warnings.filterwarnings(action='ignore') # Suppress warnings where appropriate
import xarray as xr
import numpy as np
import regionmask
import pyogrio
from joblib import Parallel, delayed

# Load shapefile and initialize region mask
countries = gpd.read_file(r"/home/pouria/git/water-institute/data/basins_charactristics/input/shapefiles/14 AllPoint SubWatershed/14_SubWatershed_All_4040217_Merged.shp")
countries = countries.to_crs(4326) # Ensure it's in WGS84 geographic CRS
output_excel_path = r"/home/pouria/git/water-institute/data/basins_charactristics/output/excel/14/location.xlsx"
indexes = list(range(len(countries)))
countries_mask_poly = regionmask.Regions(
    name='Name',
    numbers=indexes,
    names=countries['Point_ID'].astype(str).values, # Convert IDs to string for regionmask attributes
    abbrevs=countries['Point_ID'].astype(str).values, # Convert IDs to string for regionmask attributes
    outlines=list(countries.geometry.values),
    overlap=False # Crucial to avoid MemoryError by forcing a 2D mask
)



# Function to process a single country/basin (main loop unit for parallelization)
def process_country(ID_COUNTRY, countries):
    """
    Processes a single country/basin to extract its mean, 5th, and 95th percentile DEM values.

    Args:
        ID_COUNTRY (int): The integer index of the current basin in the countries GeoDataFrame.
        mask (xarray.DataArray): The full 2D mask generated by regionmask.
        lat (numpy.ndarray): Latitude coordinates from the mask.
        lon (numpy.ndarray): Longitude coordinates from the mask.
        countries (geopandas.GeoDataFrame): The GeoDataFrame containing basin polygons.
        dem_ds (xarray.Dataset): The full DEM dataset.

    Returns:
        dict: A dictionary containing the basin's mean DEM, 5th percentile, 95th percentile,
              centroid lat/lon, and basin code.
    """
    country_code = str(countries['Point_ID'][ID_COUNTRY]) # Get the string ID for the basin
    
    # Get the specific basin's geometry for centroid calculation
    basin_geometry = countries.geometry.iloc[ID_COUNTRY]
    
    # --- Centroid calculation with robust re-projection ---
    projected_crs_epsg = 32639 # Example: UTM Zone 39N. Adjust if your basins are in another zone.
    
    lon1, lat1 = np.nan, np.nan # Initialize as NaN in case of failure

    # Prioritize valid geometries and re-projection for accurate centroids
    if basin_geometry.is_valid:
        try:
            # Reproject to a projected CRS (meters) for accurate centroid calculation
            basin_geometry_proj = basin_geometry.to_crs(epsg=projected_crs_epsg)
            centroid_proj = basin_geometry_proj.centroid
            
            # Convert the centroid point back to the geographic CRS (WGS84)
            centroid_4326 = centroid_proj.to_crs(epsg=4326)
            lon1, lat1 = centroid_4326.x, centroid_4326.y
        except Exception as e:
            # This catch is for issues during re-projection itself, even if geometry is valid
            print(f"Warning: Reprojection to EPSG:{projected_crs_epsg} failed for basin {ID_COUNTRY} ({country_code}). Error: {e}")
            print("         Falling back to geographic centroid calculation (may be inaccurate and show UserWarning).")
            # Fallback to geographic centroid if re-projection fails
            # This will trigger the UserWarning if it's not projected
            centroid_geo = basin_geometry.centroid 
            lon1, lat1 = centroid_geo.x, centroid_geo.y
    else:
        print(f"Warning: Basin ID {ID_COUNTRY} ({country_code}) has an invalid geometry. Cannot calculate accurate centroid.")
        print("         Falling back to geographic centroid calculation (may be inaccurate and show UserWarning).")
        # Fallback to geographic centroid if geometry is invalid
        # This will trigger the UserWarning if it's not projected
        centroid_geo = basin_geometry.centroid
        lon1, lat1 = centroid_geo.x, centroid_geo.y
    # --- End of Centroid Fix ---
    
    # Process the DEM data for this specific basin
    c = {
        "Point_ID": country_code,
        "lat": lat1,
        "lon": lon1
    }
    return c

# Use Joblib to parallelize country processing
print("Starting parallel processing...")
results = Parallel(n_jobs=8, verbose=5)(
    delayed(process_country)(ID_COUNTRY, countries) 
    for ID_COUNTRY in indexes
)
print("Parallel processing complete.")

# Convert the list of dictionaries into a Pandas DataFrame
cc = pd.DataFrame(results)

# --- REORDER COLUMNS TO PUT 'Point_ID' FIRST ---
# Get all current column names
all_columns = cc.columns.tolist()

# Remove 'Point_ID' from its current position (it's guaranteed to be there)
all_columns.remove('Point_ID')

# Create the new desired order: 'Point_ID' first, then the rest
new_column_order = ['Point_ID'] + all_columns

# Reindex the DataFrame to apply the new order
cc = cc[new_column_order]
# --- END REORDERING ---

# Save the results to an Excel file

os.makedirs(os.path.dirname(output_excel_path), exist_ok=True) # Ensure output directory exists
cc.to_excel(output_excel_path, index=False) # index=False prevents writing the DataFrame index to Excel
print(f"Results saved to: {output_excel_path}")