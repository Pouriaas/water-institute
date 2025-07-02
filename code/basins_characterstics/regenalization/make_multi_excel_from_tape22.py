#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Jul  2 14:05:17 2025

@author: pouria
"""

import os
import re
from pathlib import Path
import pandas as pd
from tqdm import tqdm

# === CONFIGURATION ===
output_base_dir = Path("C:/Users/DrBZahraei/Desktop/regenalization/output")
output_excel_dir = Path("C:/Users/DrBZahraei/Desktop/regenalization/output/results")
output_excel_dir.mkdir(exist_ok=True)

def extract_last_hydrograph(tape22_path):
    """
    Extracts the last hydrograph (after final 3R) from a TAPE22 file.
    Returns a list of floats.
    """
    try:
        with open(tape22_path, "r") as f:
            lines = f.readlines()

        # Find last index of a line starting with '3R'
        last_3r_index = max(i for i, line in enumerate(lines) if line.strip().startswith("3R"))
        values = []

        for line in lines[last_3r_index+1:]:
            if not line.strip():
                break
            line_values = re.findall(r"[-+]?\d*\.\d+|\d+", line)
            values.extend([float(v) for v in line_values])

        return values

    except Exception as e:
        print(f"⚠️ Failed to read {tape22_path}: {e}")
        return None

# === MAIN LOOP ===
station_folders = [f for f in output_base_dir.iterdir() if f.is_dir() and f.name.isdigit()]
print(f"📁 Found {len(station_folders)} station folders.")

for station_folder in tqdm(station_folders, desc="📊 Processing stations"):
    combination_folders=os.listdir(station_folder)
    
    pathtoout=station_folder / combination_folders[0] / "out"

    folders = [f for f in os.listdir(pathtoout) if os.path.isdir(os.path.join(pathtoout, f))]
    
    excel_writer = pd.ExcelWriter(output_excel_dir / f"{station_folder.name}.xlsx", engine='openpyxl')

    # Find all event folders inside combination folders
      # {event_name: {combination_name: hydrograph_list}}

    for event_folder in folders:
        event_data = {}
        for combination_folder in combination_folders:
            
            tape22_path = station_folder / combination_folder / "out" / event_folder / "TAPE22"


            hydrograph = extract_last_hydrograph(tape22_path)
            if hydrograph is None:
                continue
            # Organize data
            event_data[combination_folder]=hydrograph
        df=pd.DataFrame(event_data)

        df.to_excel(excel_writer, sheet_name=event_folder, index=False)
        
    excel_writer.close()
    print(f"✅ Saved {station_folder.name}.xlsx")

print("\n✅ All Excel files generated in:", output_excel_dir)