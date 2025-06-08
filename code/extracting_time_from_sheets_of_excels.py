#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Jun  8 09:07:44 2025

@author: pouria
"""

import pandas as pd
import os

# Set your folder containing the Excel files
input_folder = '/home/pouria/git/water-institute/data/drmn_input/input - Copy'
output_excel = '/home/pouria/git/water-institute/data/Event_Summary_By_Basin.xlsx'

# Dictionary to hold the data for each basin
basin_data = {}

for filename in os.listdir(input_folder):
    if filename.endswith('.xlsx'):
        basin_code = os.path.splitext(filename)[0]  # Get basin code from filename
        basin_code=basin_code[:2]+basin_code[3:6]
        file_path = os.path.join(input_folder, filename)
        excel_file = pd.ExcelFile(file_path)
        
        summary = []
        
        for sheet in excel_file.sheet_names:
            df = excel_file.parse(sheet)
            if 'Date' in df.columns and not df['Date'].empty:
                # Drop NaT or blank rows in Date column
                df['Date'] = df['Date'].ffill()
                summary.append([df.loc[0,"Date"],df.loc[0,"Time"] ,df.loc[len(df)-1,"Date"],df.loc[len(df)-1,"Time"]])

        
        if summary:
            basin_data[basin_code] = pd.DataFrame(summary)

# Write to a single Excel file with one sheet per basin
with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
    for basin_code, df in basin_data.items():
        df.to_excel(writer, sheet_name=basin_code, index=False,header=False)

print(f"Summary file created: {output_excel}")