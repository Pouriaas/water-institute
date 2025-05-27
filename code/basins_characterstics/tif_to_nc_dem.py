import rioxarray as rxr
import xarray as xr

# Input and output paths
input_tif = r'/home/pouria/git/water-institute/data/basins_charactristics/input/12 DEM-4040118/12 DEM-4040118/12-Talesh-Anzali-DEM-4040118.tif'
output_nc = '/home/pouria/git/water-institute/data/basins_charactristics/output/dem_nc/12.nc'

# Read and auto-mask NoData values
dem = rxr.open_rasterio(input_tif, masked=True).squeeze()

# Reproject to WGS84
dem_wgs84 = dem.rio.reproject("EPSG:4326")

# Rename dimensions and coordinates
dem_wgs84 = dem_wgs84.rename({
    "x": "lon",
    "y": "lat"
})
dem_wgs84 = dem_wgs84.assign_coords({
    "longitude": dem_wgs84.lon,
    "latitude": dem_wgs84.lat
})

# Set metadata
dem_wgs84.name = "elevation"
dem_wgs84.attrs["units"] = "meters"
dem_wgs84.attrs["description"] = "DEM elevation data reprojected to WGS84"
dem_wgs84.longitude.attrs["units"] = "degrees_east"
dem_wgs84.latitude.attrs["units"] = "degrees_north"

# Export to NetCDF
dem_wgs84.to_netcdf(output_nc)

print(f"✅ DEM converted and saved to NetCDF with latitude/longitude coordinates: {output_nc}")