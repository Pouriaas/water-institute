#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat May 31 11:11:04 2025

@author: pouria
"""

import xarray as xr
import rioxarray as rxr # This registers the .rio accessor
import os

def convert_tif_to_nc(local_tif_filename, output_nc_filename="global_hydrologic_soil_group.nc"):
    """
    Converts a local TIFF file into a NetCDF file, handling large files efficiently
    using Dask for out-of-core processing.

    Args:
        local_tif_filename (str): The path to the local TIFF file.
        output_nc_filename (str): The desired name for the output NetCDF file.
    """
    print(f"Attempting to convert local TIFF file: {local_tif_filename}")

    if not os.path.exists(local_tif_filename):
        print(f"Error: The file '{local_tif_filename}' does not exist. Please ensure the path is correct.")
        return

    try:
        # Open the TIFF file using rioxarray, enabling Dask for out-of-core processing.
        # By specifying 'chunks', the data is not loaded entirely into memory,
        # but processed in smaller blocks, which is crucial for large files.
        # A chunk size of (1000, 1000) for (y, x) is a common starting point,
        # but you might need to adjust it based on your system's memory and file size.
        print(f"Opening {local_tif_filename} with rioxarray using Dask chunks...")
        ds = rxr.open_rasterio(local_tif_filename, chunks={'x': 1000, 'y': 1000}, decode_coords=False)

        # Ensure the CRS is properly set if it's not already.
        # Using .rio.write_crs() as recommended by rioxarray for future compatibility.
        if "crs" not in ds.coords and ds.rio.crs:
            ds = ds.rio.write_crs(ds.rio.crs) # Updated from set_crs to write_crs
            print(f"Set CRS to: {ds.rio.crs}")
        elif "crs" in ds.coords:
            print(f"Existing CRS: {ds.rio.crs}")
        else:
            print("No CRS found or set for the dataset.")

        # If the dataset has a 'band' dimension, you might want to select a specific band
        # or rename it for clarity. For single-band rasters, 'band' is often 1.
        if 'band' in ds.dims and ds.band.size == 1:
            ds = ds.squeeze('band') # Remove the band dimension if it's a single band
            print("Squeezed 'band' dimension as it contains only one band.")
        elif 'band' in ds.dims:
            print(f"Dataset has {ds.band.size} bands. Consider selecting a specific band if needed.")

        # Assign a meaningful name to the data variable if it's generic (e.g., 'band_data')
        # The original file has a single band representing "Hydrologic Soil Group"
        ds = ds.rename("hydrologic_soil_group")
        print("Renamed data variable to 'hydrologic_soil_group'.")

        # Add some metadata (optional but good practice)
        ds.attrs['description'] = 'Global Hydrologic Soil Group derived from TIFF'
        ds.attrs['source_file'] = local_tif_filename
        ds.attrs['created_by'] = 'Python script using xarray and rioxarray'

        # Save the xarray Dataset to a NetCDF file.
        # When saving a Dask-backed xarray dataset, .to_netcdf() will automatically
        # compute and write the data in chunks, preventing memory overflow.
        print(f"Saving dataset to {output_nc_filename}...")
        ds.to_netcdf(output_nc_filename)
        print(f"Successfully converted {local_tif_filename} to {output_nc_filename}")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    # IMPORTANT: Replace this with the actual path to your downloaded TIFF file.
    local_tif_path = r"/home/pouria/git/water-institute/data/basins_charactristics/input/Global_Hydrologic_Soil_Group_1566/data/HYSOGs250m.tif"
    convert_tif_to_nc(local_tif_path)