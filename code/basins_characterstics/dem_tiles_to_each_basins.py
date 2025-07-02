from osgeo import gdal, ogr
import os
import glob
import json

# --- Configuration ---
tiles_folder = r'/home/pouria/git/water-institute/data/basins_charactristics/input/GCN/New folder'
shapefile = '/home/pouria/git/water-institute/data/basins_charactristics/input/shapefiles/41 AllPoint SubWatershed/415_SubWatershed_AllPoint_4040226_Merged.shp'
output_folder = '/home/pouria/git/water-institute/data/basins_charactristics/output/GCN/415'
ndvi_files = glob.glob(os.path.join(tiles_folder, '*.tif'))

os.makedirs(output_folder, exist_ok=True)
gdal.UseExceptions()

# --- Open shapefile ---
shp_ds = ogr.Open(shapefile)
if shp_ds is None:
    raise RuntimeError(f'Cannot open shapefile: {shapefile}')
layer = shp_ds.GetLayer()

# --- Loop over each feature (basin) ---
for feature in layer:
    point_id = feature.GetField('Point_ID')
    geometry = feature.GetGeometryRef()
    if not geometry:
        print(f'⚠️ Skipping feature with missing geometry (Point_ID: {point_id})')
        continue

    # Check validity
    if not geometry.IsValid():
        print(f'⚠️ Geometry invalid for Point_ID={point_id}, attempting to fix...')
        fixed = geometry.Buffer(0)
        if not fixed or not fixed.IsValid():
            print(f'❌ Could not fix geometry. Skipping Point_ID={point_id}')
            continue
        geometry = fixed

    # Convert to GeoJSON
    geom_json = geometry.ExportToJson()
    geojson_dict = {
        "type": "FeatureCollection",
        "features": [{
            "type": "Feature",
            "geometry": json.loads(geom_json),
            "properties": {"Point_ID": point_id}
        }]
    }

    # Write to virtual in-memory GeoJSON
    geojson_path = '/vsimem/temp_basin.geojson'
    geojson_bytes = json.dumps(geojson_dict).encode('utf-8')
    f = gdal.VSIFOpenL(geojson_path, 'w')
    gdal.VSIFWriteL(geojson_bytes, 1, len(geojson_bytes), f)
    gdal.VSIFCloseL(f)

    # Output filename
    output_path = os.path.join(output_folder, f'ndvi_{point_id}.tif')

    # Try clipping
    try:
        gdal.Warp(
            output_path,
            ndvi_files,
            cutlineDSName=geojson_path,
            cropToCutline=True,
            dstNodata=-9999,
            format='GTiff',
            creationOptions=['COMPRESS=LZW']
        )
        print(f'✅ Saved clipped NDVI for point_id={point_id}')
    except RuntimeError as e:
        print(f'❌ Failed to clip Point_ID={point_id}: {e}')

    gdal.Unlink(geojson_path)

# Done
print('🎉 All basins processed.')
