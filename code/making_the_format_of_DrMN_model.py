import numpy as np
import pandas as pd
import os
import math
import re
from openpyxl import load_workbook

input_file = r'/home/pouria/git/water-institute/data/drmn_input/input'
output_file = r"/home/pouria/git/water-institute/data/drmn_input/output"


point_tp=pd.read_excel(r"/home/pouria/git/water-institute/data/drmn_input/12 PointTemperature/12 PointTemperature-4040128.xlsx",sheet_name="Sheet1").set_index("Point_ID")

hydro_attribute=pd.read_excel(r"/home/pouria/git/water-institute/data/drmn_input/Hydro_Attribute Table.xlsx").set_index("Code")

characteristic = pd.read_excel(r"/home/pouria/git/water-institute/data/drmn_input/12_SubWatershed_Characteristics_HydroStations/12_Talesh_Anzali_SubWatershed_Characteristics_HydroStations_4040125.xlsx",sheet_name="Sheet1", skiprows=2).set_index("Point_ID")

snow = pd.read_excel(r"/home/pouria/git/water-institute/data/drmn_input/12_SubWatershed_Characteristics_HydroStations/12_Talesh_Anzali_SubWatershed_Characteristics_HydroStations_4040125.xlsx",sheet_name="Sheet1").set_index("Point_ID")

hypsometric_folder=r"/home/pouria/git/water-institute/data/drmn_input/Hypsometric_Tables"
files_and_directory = os.listdir(input_file)

TemperatureGradientfolder=r"/home/pouria/git/water-institute/data/drmn_input/12 TemperatureGradient"
for name in files_and_directory:
    path = os.path.join(input_file, name)
    # Step 1: Load the original workbook with openpyxl
    wb = load_workbook(path, data_only=False)
    
    # Step 2: Capture the tab colors of existing sheets
    tab_colors = {sheet.title: sheet.sheet_properties.tabColor.rgb for sheet in wb.worksheets}
    
    path1 = os.path.join(output_file, name)
    
    rename=name[:2]+name[3:6]
    
    
    # Load the Excel file
    excel_file = pd.ExcelFile(path)
    hypsometric_path=os.path.join(hypsometric_folder,rename+'.xlsx')
    hypsometric=pd.read_excel(hypsometric_path)
    
    # Get all sheet names
    sheet_names = excel_file.sheet_names
    arr=np.full([9,10],np.nan)
    sim=pd.DataFrame(arr.T,columns=["Tc (hr)",	"Hydrometry Station Elevation",	"Elevation Range (m) strat","Elevation Range (m) end",		"Area (km2)",	"Zone height (m)",	"Snow melt coeff","	Freezing temp (C)",	"Basin_3rd",	"Total Area (km2)","POINT_X","POINT_Y"])
    sim.loc[0, "Basin_3rd"] = hydro_attribute.loc[int(rename), "Basin_3rd"]
    sim.loc[0, "POINT_X"] = hydro_attribute.loc[int(rename), "POINT_X"]
    sim.loc[0, "POINT_Y"] = hydro_attribute.loc[int(rename), "POINT_Y"]
    sim.loc[0, "Tc (hr)"] = characteristic.loc[int(rename), "Tc(Krp_hr)"]
    sim.loc[0, "Total Area (km2)"] = hypsometric.loc[0, "Cumulative Area (km²)"]
    sim.loc[0, "Snow melt coeff"] = snow.loc[sim.loc[0, "Basin_3rd"], "Snow melt coeff"]
    sim.loc[0, "Freezing temp (C)"] = snow.loc[sim.loc[0, "Basin_3rd"], "Freezing temp (C)"]
    TemperatureGradientpath=os.path.join(TemperatureGradientfolder,str(sim.loc[0, "Basin_3rd"])[:-2]+".xlsx")
    TemperatureGradient=pd.read_excel(TemperatureGradientpath)
    
    # Extract start and end elevations
    hypsometric[['Start', 'End']] = hypsometric['Elevation Range (m)'].str.split(' - ', expand=True).astype(float)
    
    # Sort by start elevation (just in case)
    hypsometric = hypsometric.sort_values('Start').reset_index(drop=True)
    
    # Determine group size
    num_rows = len(hypsometric)
    num_groups = min(10, num_rows)
    group_size = math.ceil(num_rows / num_groups)
    
    # Assign group number to each row
    hypsometric['Group'] = hypsometric.index // group_size
    
    # Group and merge
    grouped = hypsometric.groupby('Group').agg({
        'Start': 'min',
        'End': 'max',
        'Area (km²)': 'sum'
    }).reset_index(drop=True)
    
    sim.loc[:len(grouped),"Elevation Range (m) strat"]=grouped['Start'][:]
    sim.loc[:len(grouped),"Elevation Range (m) end"]=grouped['End'][:]
    sim.loc[:len(grouped),"Area (km2)"]=grouped['Area (km²)'][:]
    sim.loc[0,"Hydrometry Station Elevation"]=grouped.loc[0,'Start']
    sim.loc[0,"Zone height (m)"]=grouped.loc[0,'End']-grouped.loc[0,'Start']
    
    # Create a writer to save output
    with pd.ExcelWriter(path1, engine='openpyxl') as writer:
        for sheet_name in sheet_names:
            df = pd.read_excel(path, sheet_name=sheet_name)
            
            # Fill NaN in existing 'Date'
            df['Date'] = df['Date'].ffill()
            n_rows = len(df)
            
            zero=np.full(n_rows,np.nan,dtype=object)
            new_date=df['Date'].dropna().unique()
            
            # Add new columns with same name 'Date' + Temperature + temp.lapse/zone
            df['Date_new'] = zero # avoid length mismatch
            df['Temperature'] = zero
            df['temp.lapse/zone'] = zero
            
            df['Date_new'][:len(new_date)] = new_date # avoid length mismatch
            tp=[]
            for datee in new_date:
                year=datee.split("/")[0]
                datee=datee.replace("/","")
                if int(year)<3:
                    datee=str("14"+datee)
                else:
                    datee=str("13"+datee)
                tp.append(point_tp.loc[int(rename), int(datee)])
            
            
            df['Temperature'][:len(new_date)]  = tp
            month = int(str(df.loc[0, 'Date']).split("/")[1])
            formula=TemperatureGradient.iloc[month-1,1]
            match = abs(float(re.search(r'([-+]?\d*\.?\d+)\*H', formula).group(1)))
            
            df['temp.lapse/zone'][0]=match*sim.loc[0,"Zone height (m)"]
            
            
            # Rename new Date column back to 'Date' to create duplicate names
            df.rename(columns={'Date_new': 'Date'}, inplace=True)
            
            # Write to new Excel file
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        sim.to_excel(writer, sheet_name="Geometry", index=False)
    # Step 4: Reopen to reassign colors and reorder sheets
    wb = load_workbook(path1)
    
    # Reapply original tab colors
    for sheet in wb.worksheets:
        if sheet.title in tab_colors and tab_colors[sheet.title] is not None:
            sheet.sheet_properties.tabColor = tab_colors[sheet.title]
    
    # Move Geometry to the front
    geometry_sheet = wb["Geometry"]
    wb._sheets.remove(geometry_sheet)
    wb._sheets.insert(0, geometry_sheet)
    
    # Step 5: Save to original path (or new path)
    wb.save(path1)
