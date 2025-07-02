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
import rioxarray 


output_excel_path = r"/home/pouria/git/water-institute/data/basins_charactristics/output/dem_excel/soiltype_12.xlsx"
# --- IMPORTANT CORRECTION HERE ---
# The 'countries' variable MUST be your shapefile (vector data) defining the basins.
# Please ensure this path points to your actual basin shapefile.
countries = gpd.read_file(r"/home/pouria/git/water-institute/data/basins_charactristics/input/12 AllPoint SubWatershed/12_SubWatershed_All_4040123_Merged.shp")
countries = countries.to_crs(4326) # Ensure it's in WGS84 geographic CRS

# Load the raster dataset (e.g., Soil Type, DEM) using rioxarray for the TIFF file.
# This path should point to your 'HYSOGs250m.tif' file as per your error message.
try:
    # --- IMPORTANT CHANGE: Added chunks='auto' to enable Dask-backed lazy loading ---
    dem_ds = rioxarray.open_rasterio(
        r"/home/pouria/git/water-institute/data/basins_charactristics/input/Global_Hydrologic_Soil_Group_1566/data/HYSOGs250m.tif",
        decode_times=False,
        chunks='auto' # This tells rioxarray/xarray to load the data in chunks using Dask
    )
except Exception as e:
    print(f"Error opening TIFF file: {e}")
    print("Please ensure the TIFF file exists at the specified path and is not corrupted.")
    print("You can try opening the TIFF file in a GIS software (e.g., QGIS) to verify its integrity.")
    exit() # Exit the script if the raster file cannot be opened

# If the TIFF has a single band, remove the 'band' dimension for easier processing
if 'band' in dem_ds.dims and dem_ds.sizes['band'] == 1:
    dem_ds = dem_ds.squeeze('band', drop=True)

# --- NEW FIX: Explicitly convert nodata values (255) to NaN and cast to float dtype ---
if dem_ds.rio.nodata is not None:
    nodata_value = dem_ds.rio.nodata
    print(f"Detected nodata value: {nodata_value}. Converting to NaN...")
    # Convert to float to allow NaN values
    dem_ds = dem_ds.astype(np.float32) # Use float32 for memory efficiency unless higher precision is needed
    # Replace the nodata value with NaN
    dem_ds = dem_ds.where(dem_ds != nodata_value, np.nan) # Explicitly set condition to NaN
    # Ensure rioxarray knows the nodata is now NaN (optional, but good practice)
    dem_ds.rio.write_nodata(np.nan, inplace=True)
else:
    print("No explicit nodata value found in TIFF metadata or already converted.")
    # If no nodata is explicitly set but dtype is int, consider converting to float anyway
    if np.issubdtype(dem_ds.dtype, np.integer):
        print("Data type is integer, converting to float for potential NaN handling.")
        dem_ds = dem_ds.astype(np.float32)
print(f"DEM DataArray dtype after nodata handling: {dem_ds.dtype}")
# --- END NEW FIX ---

# --- NEW ADDITION: Print the range of values in the initial TIFF file ---
# print("--- Initial TIFF File Value Range ---")
# .min().compute() and .max().compute() will use Dask to find the min/max across the entire file
# min_val = dem_ds.min(skipna=True).compute().item() # .item() converts 0-d array to scalar
# max_val = dem_ds.max(skipna=True).compute().item()
# print(f"Minimum value in original TIFF (excluding NaN): {min_val}")
# print(f"Maximum value in original TIFF (excluding NaN): {max_val}")
# print("-------------------------------------")
# --- END NEW ADDITION ---

# --- NEW OPTIMIZATION: Perform an initial rectangular clip of the large TIFF ---
# This step should happen BEFORE renaming 'x'/'y' to 'lon'/'lat'
# as rioxarray's clip_box expects 'x' and 'y' by default.
print("Performing initial rectangular clip of the raster to the basins' extent...")
# Get the total bounding box of all basins
minx, miny, maxx, maxy = countries.total_bounds

# Clip the raster to this bounding box
# This creates a smaller Dask-backed DataArray for subsequent processing
dem_ds_subset = dem_ds.rio.clip_box(minx, miny, maxx, maxy, crs=countries.crs)
print("Initial rectangular clip complete.")

# --- DEBUGGING: Print information about the subsetted raster ---
print(f"dem_ds_subset CRS: {dem_ds_subset.rio.crs}")
print(f"dem_ds_subset bounds: {dem_ds_subset.rio.bounds()}")
print(f"dem_ds_subset nodata: {dem_ds_subset.rio.nodata}") # Should now be nan
print(f"dem_ds_subset dtype: {dem_ds_subset.dtype}") # Should now be float32
print(f"dem_ds_subset shape: {dem_ds_subset.shape}")
# --- FIX: Call .compute() on the DataArray before accessing .values ---
print(f"dem_ds_subset has valid data (first 10 computed): {dem_ds_subset.compute().values.flatten()[:10]}")
# --- END DEBUGGING ---

# --- IMPORTANT CHANGE START: Ensure 'lat' and 'lon' coordinates are present in the SUBSET ---
# This block is now applied to dem_ds_subset AFTER the initial clip.
# Check if 'y' and 'x' are present and rename them to 'lat' and 'lon' for consistency
# This ensures regionmask and other xarray operations work correctly later.
if 'y' in dem_ds_subset.coords and 'x' in dem_ds_subset.coords:
    dem_ds_subset = dem_ds_subset.rename({'y': 'lat', 'x': 'lon'})
elif 'lat' not in dem_ds_subset.coords or 'lon' not in dem_ds_subset.coords:
    raise ValueError("Could not find 'lat'/'lon' or 'y'/'x' coordinates in the raster data after clipping. Please check your TIFF file's coordinate system.")
# --- IMPORTANT CHANGE END ---


# Function to process a single country/basin (main loop unit for parallelization)
def process_country(ID_COUNTRY_idx, countries_gdf, full_dem_ds_subset):
    """
    Processes a single country/basin to extract its raster statistics.
    This function now handles clipping and masking per basin to optimize memory.

    Args:
        ID_COUNTRY_idx (int): The integer index of the current basin in the countries GeoDataFrame.
        countries_gdf (geopandas.GeoDataFrame): The GeoDataFrame containing basin polygons.
        full_dem_ds_subset (xarray.DataArray): The Dask-backed raster DataArray,
                                                already clipped to the overall basin extent
                                                and with 'lat'/'lon' dimensions.

    Returns:
        dict: A dictionary containing the basin's raster statistics,
              centroid lat/lon, and basin code.
    """
    country_code = str(countries_gdf['Point_ID'][ID_COUNTRY_idx]) # Get the string ID for the basin
    
    # Get the specific basin's geometry for centroid calculation
    basin_geometry = countries_gdf.geometry.iloc[ID_COUNTRY_idx]
    
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
            print(f"Warning: Reprojection to EPSG:{projected_crs_epsg} failed for basin {ID_COUNTRY_idx} ({country_code}). Error: {e}")
            print("         Falling back to geographic centroid calculation (may be inaccurate and show UserWarning).")
            centroid_geo = basin_geometry.centroid 
            lon1, lat1 = centroid_geo.x, centroid_geo.y
    else:
        print(f"Warning: Basin ID {ID_COUNTRY_idx} ({country_code}) has an invalid geometry. Cannot calculate accurate centroid.")
        print("         Falling back to geographic centroid calculation (may be inaccurate and show UserWarning).")
        centroid_geo = basin_geometry.centroid
        lon1, lat1 = centroid_geo.x, centroid_geo.y
    # --- End of Centroid Fix ---
    
    # --- IMPORTANT MEMORY OPTIMIZATION START ---
    p5_val, p50_val, p95_val = np.nan, np.nan, np.nan # Initialize values

    try:
        # 1. Clip the full_dem_ds_subset to the bounding box of the current basin
        # This operation is Dask-backed and only defines the computation graph.
        # `drop=True` ensures that areas outside the clip are dropped, reducing array size.
        # This clip is now on an already smaller subset!
        clipped_dem_ds = full_dem_ds_subset.rio.clip([basin_geometry], countries_gdf.crs, drop=True)

        # --- DEBUGGING: Print information about the clipped raster for this basin ---
        # Removed extensive per-basin debugging prints to reduce console clutter for successful runs.
        # Re-enable if needing to debug NaNs again.
        # print(f"Basin {country_code} (ID: {ID_COUNTRY_idx}) clipped_dem_ds size: {clipped_dem_ds.size}")
        # if clipped_dem_ds.size > 0:
        #     print(f"Basin {country_code} clipped_dem_ds nodata: {clipped_dem_ds.rio.nodata}")
        #     print(f"Basin {country_code} clipped_dem_ds CRS: {clipped_dem_ds.rio.crs}")
        #     print(f"Basin {country_code} clipped_dem_ds bounds: {clipped_dem_ds.rio.bounds()}")
        #     clipped_values_sample = clipped_dem_ds.compute().values.flatten()
        #     print(f"Basin {country_code} clipped_dem_ds values (sample, non-NaN): {clipped_values_sample[~np.isnan(clipped_values_sample)][:10]}")
        # --- END DEBUGGING ---

        if clipped_dem_ds.size == 0:
            print(f"Warning: Basin ID {ID_COUNTRY_idx} ({country_code}) has no overlap with raster data after clipping. Setting statistics to NaN.")
        else:
            # 2. Create a temporary regionmask.Regions object for *only this basin*
            # This is a small, localized mask definition.
            single_basin_regions = regionmask.Regions(
                numbers=[ID_COUNTRY_idx], # Use the basin's index as its number for masking
                names=[country_code],
                abbrevs=[country_code],
                outlines=[basin_geometry],
                overlap=False
            )

            # 3. Mask the clipped DEM with the single basin geometry.
            # This returns a Dask-backed DataArray.
            basin_mask_dataarray = single_basin_regions.mask(clipped_dem_ds)

            # 4. Apply the mask condition: set cells outside the basin to NaN
            # This also returns a Dask-backed DataArray.
            dem_data_slice = clipped_dem_ds.where(basin_mask_dataarray == ID_COUNTRY_idx)

            # --- DEBUGGING: Print information about the final data slice for this basin ---
            dem_data_slice_computed = dem_data_slice.compute()
            non_nan_count = np.count_nonzero(~np.isnan(dem_data_slice_computed.values))
            # print(f"Basin {country_code} dem_data_slice_computed non-NaN count: {non_nan_count}")
            # if non_nan_count > 0:
            #     print(f"Basin {country_code} dem_data_slice_computed values (sample, non-NaN): {dem_data_slice_computed.values[~np.isnan(dem_data_slice_computed.values)][:10]}")
            # else:
            #     print(f"DEBUG: All values in dem_data_slice_computed for basin {country_code} are NaN.")
            # --- END DEBUGGING ---

            # 5. Compute statistics on the Dask-backed DataArray.
            # The .compute() method here triggers the actual loading and calculation
            # only for the relevant, small chunk of data for the current basin.
            if dem_data_slice.size > 0 and non_nan_count > 0: # Check for non-NaN count before computing percentiles
                percentiles_data = dem_data_slice.quantile([0.05, 0.5, 0.95], skipna=True).compute().values
                p5_val = percentiles_data[0]
                p50_val = percentiles_data[1]
                p95_val = percentiles_data[2]
            else:
                print(f"Warning: No valid raster data found for Basin ID {ID_COUNTRY_idx} ({country_code}) after masking. Setting statistics to NaN.")

    except Exception as e:
        print(f"Error processing raster data for basin {ID_COUNTRY_idx} ({country_code}): {e}")
        print("Setting statistics to NaN for this basin due to an error during raster processing.")
    # --- IMPORTANT MEMORY OPTIMIZATION END ---

    # Prepare the result dictionary with updated names
    c = {
        "lat": lat1,
        "lon": lon1,
        "soiltype_p5": p5_val,
        "soiltype_p50": p50_val,
        "soiltype_p95": p95_val,
    }
    # --- FIX: Add 'Point_ID' to the dictionary before returning it ---
    c.update({"Point_ID": country_code})
    return c

# Use Joblib to parallelize country processing
print("Starting parallel processing...")
# We pass the pre-clipped dem_ds_subset to each parallel process.
results = Parallel(n_jobs=11, verbose=5)(
    delayed(process_country)(ID_COUNTRY_idx, countries, dem_ds_subset) 
    for ID_COUNTRY_idx in range(len(countries)) # Iterate through the indices of the countries GeoDataFrame
)
print("Parallel processing complete.")

# Convert the list of dictionaries into a Pandas DataFrame
cc = pd.DataFrame(results)

# --- REORDER COLUMNS TO PUT 'Point_ID' FIRST ---
all_columns = cc.columns.tolist()
if 'Point_ID' in all_columns: # Ensure 'Point_ID' exists before trying to remove/reorder
    all_columns.remove('Point_ID')
new_column_order = ['Point_ID'] + all_columns
cc = cc[new_column_order]
# --- END REORDERING ---

# Save the results to an Excel file with the updated name

os.makedirs(os.path.dirname(output_excel_path), exist_ok=True) # Ensure output directory exists
cc.to_excel(output_excel_path, index=False) # index=False prevents writing the DataFrame index to Excel
print(f"Results saved to: {output_excel_path}")