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
countries = gpd.read_file(r"/home/pouria/git/water-institute/data/basins_charactristics/input/12 AllPoint SubWatershed/12_SubWatershed_All_4040123_Merged.shp")
countries = countries.to_crs(4326) # Ensure it's in WGS84 geographic CRS

indexes = list(range(len(countries)))
countries_mask_poly = regionmask.Regions(
    name='Name',
    numbers=indexes,
    names=countries['Point_ID'].astype(str).values, # Convert IDs to string for regionmask attributes
    abbrevs=countries['Point_ID'].astype(str).values, # Convert IDs to string for regionmask attributes
    outlines=list(countries.geometry.values),
    overlap=False # Crucial to avoid MemoryError by forcing a 2D mask
)

# Load the DEM dataset once globally using open_dataset for a single file
dem_ds = xr.open_dataset(r"/home/pouria/git/water-institute/data/basins_charactristics/output/dem_nc/12.nc", decode_times=False)

# Create the full mask once
# This mask will have integer IDs where basins are, and NaN elsewhere
mask = countries_mask_poly.mask(dem_ds)

# Extract latitude and longitude arrays from the mask for slicing
lat = mask.lat.values
lon = mask.lon.values

# Helper function to process each basin's DEM data
def process_dataset(dem_data_obj, mask_array, lats, lons, ID_COUNTRY, lat1, lon1):
    """
    Extracts and computes the mean, 5th, and 95th percentile DEM values for a specific basin.

    Args:
        dem_data_obj (xarray.Dataset): The full DEM dataset.
        mask_array (xarray.DataArray): The full 2D mask generated by regionmask.
        lats (numpy.ndarray): Latitude coordinates from the mask.
        lons (numpy.ndarray): Longitude coordinates from the mask.
        ID_COUNTRY (int): The integer ID of the current basin.
        lat1 (float): Latitude of the basin's centroid.
        lon1 (float): Longitude of the basin's centroid.

    Returns:
        dict: A dictionary containing the mean DEM, 5th percentile, 95th percentile,
              centroid lat/lon, and basin ID.
    """
    
    # Find the indices (row, col) in the global mask where the current basin's ID is present
    y_indices, x_indices = np.where(mask_array.values == ID_COUNTRY)
    
    # Check if the basin actually covers any grid cells in the mask
    if len(y_indices) > 0 and len(x_indices) > 0:
        # Determine the bounding box (min/max *indices*) for the current basin
        min_lat_idx, max_lat_idx = y_indices.min(), y_indices.max()
        min_lon_idx, max_lon_idx = x_indices.min(), x_indices.max()
        
        # Use isel (integer-based selection) for direct slicing by index
        # Always slice from min_index to max_index + 1 (standard Python slicing)
        dem_chunk = dem_data_obj.isel(
            lat=slice(min_lat_idx, max_lat_idx + 1),
            lon=slice(min_lon_idx, max_lon_idx + 1)
        ).compute() # Load this chunk into memory for efficient processing
        
        # Select the corresponding spatial chunk from the full mask array using isel too
        # This ensures the sub_mask perfectly matches the dem_chunk's dimensions and coordinates
        sub_mask = mask_array.isel(
            lat=slice(min_lat_idx, max_lat_idx + 1),
            lon=slice(min_lon_idx, max_lon_idx + 1)
        )
        
        # Apply the mask condition: set cells outside the basin (or not matching ID_COUNTRY) to NaN
        out_sel1 = dem_chunk.where(sub_mask == ID_COUNTRY)
        
    else:
        # Fallback: If a basin polygon does not spatially overlap with any grid cell in the mask
        print(f"Warning: Basin ID {ID_COUNTRY} at ({lat1:.4f},{lon1:.4f}) not found in mask "
              f"or covers no grid cells. Falling back to nearest point DEM value.")
        out_sel1 = dem_data_obj.sel(
            lat=lat1, lon=lon1, method="nearest"
        ).compute()
    
    # Explicitly define the name of the DEM variable in your NetCDF file
    var_name = 'elevation' 
    
    # Check if the expected variable exists in the selected data
    if var_name not in out_sel1.data_vars:
        raise ValueError(f"Expected data variable '{var_name}' not found in the DEM data for Basin ID {ID_COUNTRY}. "
                         f"Available variables: {list(out_sel1.data_vars.keys())}")
    
    # --- Calculate Mean, 5th, and 95th Percentile ---
    dem_data_slice = out_sel1[var_name] # Get the DataArray for elevation
    
    # Calculate mean, skipping NaNs
    mean_dem = dem_data_slice.mean(skipna=True).values
    min_dem = dem_data_slice.min().item()
    max_dem = dem_data_slice.max().item()

    # --- End Calculations ---

    # Prepare the result dictionary
    c = {

        "lat": lat1,
        "lon": lon1,
        "dem_min": min_dem,
        "dem_mean": mean_dem, # Renamed to be more explicit
        "dem_max": max_dem,
    }

    return c

# Function to process a single country/basin (main loop unit for parallelization)
def process_country(ID_COUNTRY, mask, lat, lon, countries, dem_ds):
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
    result_dict = process_dataset(dem_ds, mask, lat, lon, ID_COUNTRY, lat1, lon1)
    result_dict.update({"Point_ID": country_code}) # Add the basin's string ID to the result
    return result_dict

# Use Joblib to parallelize country processing
print("Starting parallel processing...")
results = Parallel(n_jobs=11, verbose=5)(
    delayed(process_country)(ID_COUNTRY, mask, lat, lon, countries, dem_ds) 
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
output_excel_path = r"/home/pouria/git/water-institute/data/basins_charactristics/output/dem_excel/dem_12.xlsx"
os.makedirs(os.path.dirname(output_excel_path), exist_ok=True) # Ensure output directory exists
cc.to_excel(output_excel_path, index=False) # index=False prevents writing the DataFrame index to Excel
print(f"Results saved to: {output_excel_path}")