// --- Configuration ---
// Replace this with the Asset ID of your uploaded shapefile.
var basins = ee.FeatureCollection(table); // <-- IMPORTANT: Update this!

// Define the study period.
var startYear = 2011;
var endYear = 2024;

// Define the output file name for the CSV export.
var outputFileName = 'Basin_NDVI_Statistics_2011_2024';

// --- Load and Process NDVI Data ---

// Load the Landsat Collection 2 Tier 1 Level 2 Annual NDVI Composite.
// Filter by year and calculate the mean NDVI for each pixel over the entire period.
var annualNdvlCollection = ee.ImageCollection('LANDSAT/COMPOSITES/C02/T1_L2_ANNUAL_NDVI')
    .filter(ee.Filter.calendarRange(startYear, endYear, 'year'));

// Calculate the mean NDVI for each pixel over the entire period (2011-2024).
// This creates a single image where each pixel value is its temporal mean NDVI.
var meanNdvlImage = annualNdvlCollection.mean();

// --- Calculate Zonal Statistics for Each Basin ---

// Define the reducers outside the function for clarity.
var meanReducer = ee.Reducer.mean();
var percentileReducer = ee.Reducer.percentile([5, 95]);

// Combine the reducers.
// We set sharedInputs to true because both reducers operate on the same set of pixels.
// The second argument (null) is for an output prefix, which isn't strictly needed here
// because the percentile reducer's outputs are already clearly named (e.g., _p5, _p95).
var combinedReducer = meanReducer.combine(percentileReducer, null, true);


// Function to calculate statistics for a single basin.
var calculateBasinStats = function(feature) {
  // Clip the mean NDVI image to the current basin boundary.
  var clippedNdvl = meanNdvlImage.clip(feature.geometry());

  // Calculate the mean, 5th percentile, and 95th percentile of NDVI within the basin.
  var stats = clippedNdvl.reduceRegion({
    reducer: combinedReducer, // Use the pre-defined combined reducer
    geometry: feature.geometry(),
    scale: 30, // Landsat resolution in meters. Adjust if your basins are very small.
    maxPixels: 1e10 // Increase if you encounter 'Too many pixels' error.
  });

  // Get the name/ID of the basin. You might need to adjust 'name_field'
  // to match the actual property name in your shapefile that contains the basin ID/name.
var basinId = feature.get('Point_ID');  // <-- Use 'Point_ID' instead of 'NAME'

  // Return a new feature with the original properties and the calculated statistics.
  return feature.set({
    'Basin_ID': basinId,
    'Mean_NDVI_2011_2024': stats.get('NDVI_mean'),
    'Percentile_5_NDVI': stats.get('NDVI_p5'),
    'Percentile_95_NDVI': stats.get('NDVI_p95')
  });
};

// Map the statistics calculation function over all basins.
var basinStatistics = basins.map(calculateBasinStats);

// --- Export Results ---

// Select the properties you want in the CSV.
var propertiesToExport = [
  'Basin_ID',
  'Mean_NDVI_2011_2024',
  'Percentile_5_NDVI',
  'Percentile_95_NDVI'
];

// Export the FeatureCollection to a CSV file in your Google Drive.
Export.table.toDrive({
  collection: basinStatistics,
  description: outputFileName,
  folder: 'GEE_Exports', // This folder will be created in your Google Drive if it doesn't exist.
  fileNamePrefix: outputFileName,
  fileFormat: 'CSV',
  selectors: propertiesToExport
});

// --- Visualization (Optional) ---
// Add the mean NDVI image to the map.
Map.addLayer(meanNdvlImage, {min: 0, max: 0.8, palette: ['red', 'yellow', 'green']}, 'Mean NDVI (2011-2024)');
Map.centerObject(basins, 8); // Center the map on your basins.
Map.addLayer(basins, {color: 'blue'}, 'Basins');

print('Processing complete. Check the "Tasks" tab to run the export to Google Drive.');
print('Remember to adjust the \'NAME\' field in the code to match your basin ID property.');