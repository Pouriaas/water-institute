// --- Configuration ---
// Replace this with the Asset ID of your uploaded shapefile.
var basins = ee.FeatureCollection('YOUR_SHAPEFILE_ASSET_ID'); // <-- IMPORTANT: Update this!

// Define the study period.
var startYear = 2011;
var endYear = 2024;

// Define the output file name for the CSV export.
var outputFileName = 'Basin_Landsat_EVI_Statistics_2011_2024';

// --- Load and Process EVI Data (Landsat Annual Composite) ---

// Load the Landsat Collection 2 Tier 1 Level 2 Annual EVI Composite.
// Filter by year and select the 'EVI' band.
var annualEviCollection = ee.ImageCollection('LANDSAT/COMPOSITES/C02/T1_L2_ANNUAL_EVI')
    .filter(ee.Filter.calendarRange(startYear, endYear, 'year'))
    .select('EVI'); // Select the EVI band

// Calculate the mean EVI for each pixel over the entire period (2011-2024).
// This creates a single image where each pixel value is its temporal mean EVI.
// No scaling needed, as this composite provides EVI directly in [-1, 1] range.
var meanEviImage = annualEviCollection.mean();

// --- Calculate Zonal Statistics for Each Basin ---

// Define the reducers for mean, 5th, and 95th percentile.
var meanReducer = ee.Reducer.mean();
var percentileReducer = ee.Reducer.percentile([5, 95]);

// Combine the reducers.
var combinedReducer = meanReducer.combine(percentileReducer, null, true);

// Function to calculate statistics for a single basin.
var calculateBasinStats = function(feature) {
  // Clip the mean EVI image to the current basin boundary.
  var clippedEvi = meanEviImage.clip(feature.geometry());

  // Calculate the mean, 5th percentile, and 95th percentile of EVI within the basin.
  var stats = clippedEvi.reduceRegion({
    reducer: combinedReducer,
    geometry: feature.geometry(),
    scale: 30, // Landsat resolution is 30 meters.
    maxPixels: 1e10 // Increase if you encounter 'Too many pixels' error.
  });

  // Get the ID of the basin using 'Point_ID' field.
  var basinId = feature.get('Point_ID'); // <--- !!! IMPORTANT: Using 'Point_ID' as requested !!!

  // Return a new feature with the original properties and the calculated statistics.
  return feature.set({
    'Basin_ID': basinId,
    'Mean_EVI_2011_2024': stats.get('EVI_mean'),
    'Percentile_5_EVI': stats.get('EVI_p5'),
    'Percentile_95_EVI': stats.get('EVI_p95')
  });
};

// Map the statistics calculation function over all basins.
var basinStatistics = basins.map(calculateBasinStats);

// --- Export Results ---

// Select the properties you want in the CSV.
var propertiesToExport = [
  'Basin_ID',
  'Mean_EVI_2011_2024',
  'Percentile_5_EVI',
  'Percentile_95_EVI'
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
// Define EVI visualization parameters (EVI typically ranges from -1 to 1).
var eviVis = {
  min: 0, // EVI values usually start around 0 for bare soil, up to 1 for dense vegetation
  max: 0.8,
  palette: [
    'FFFFFF', 'CE7E45', 'DF923D', 'F1B555', 'FCD163', '99B718', '74A901',
    '66A000', '529400', '3E8601', '207401', '056201', '004C00', '023B01',
    '012E01', '011D01', '011301'
  ],
};
Map.addLayer(meanEviImage, eviVis, 'Mean EVI (2011-2024)');
Map.centerObject(basins, 8); // Center the map on your basins.
Map.addLayer(basins, {color: 'blue'}, 'Basins');

print('Processing complete. Check the "Tasks" tab to run the export to Google Drive.');
print('Using Landsat Collection 2 Tier 1 Level 2 Annual EVI Composite (30m resolution) and \'Point_ID\' from your shapefile.');