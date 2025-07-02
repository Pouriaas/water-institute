import os
import re
import pandas as pd
import openpyxl

# Base folder path
base_path = r"D:\project\code new\temp1"
subfolders = [f for f in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, f))]

# Loop through all subfolders
for folder_name in subfolders:
    folder_path = os.path.join(base_path, folder_name)
    print(f"\n[•] Processing folder: {folder_path}")

    # Locate input Excel file
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsm')]
    if not excel_files:
        print("[!] No .xlsm file found. Skipping.")
        continue
    input_excel_path = os.path.join(folder_path, excel_files[0])
    xls = pd.ExcelFile(input_excel_path)

    # Identify valid sheets
    sheet_names = [s for s in xls.sheet_names if s != "Geometry"]

    # Locate PSO file
    plot_candidates = [f for f in os.listdir(folder_path) if f.startswith("PSO_plot04")]
    if not plot_candidates:
        print("[!] No PSO_plot04 file found. Skipping.")
        continue
    plot_path = os.path.join(folder_path, plot_candidates[0])
    print(f"PSO file found: {plot_path}")

    # Read PSO file
    with open(plot_path, "r", encoding='utf-8', errors='ignore') as file:
        lines = file.readlines()

    # Prepare output writer
    output_path = os.path.join(folder_path, "output_plot04.xlsx")
    writer = pd.ExcelWriter(output_path, engine='openpyxl')
    data_frames = {}

    # Step 1: Build Observed column
    for sheet_name in sheet_names:
        df = xls.parse(sheet_name, header=None)
        df = df.iloc[1:].reset_index(drop=True)
        header_row = df.iloc[0]

        q_col_candidates = header_row[header_row.astype(str).str.strip().str.contains("Q", case=False, na=False)]
        if q_col_candidates.empty:
            print(f"[!] Q column not found in sheet '{sheet_name}'. Skipping.")
            continue

        q_col_index = q_col_candidates.index[0]
        observed_data = df[q_col_index].iloc[1:].reset_index(drop=True)
        output_df = pd.DataFrame({"Observed": observed_data})
        data_frames[sheet_name] = output_df

    # Step 2: Handle ZONE blocks
    current_zone = None
    current_data = []

    def process_zone(zone, data_lines):
        date_str = zone['date'].replace("/", ",")
        if date_str not in data_frames:
            return

        df_sheet = xls.parse(date_str, header=None)
        df_sheet = df_sheet.iloc[1:].reset_index(drop=True)
        header_row = df_sheet.iloc[0]

        strtq_candidates = header_row[header_row.astype(str).str.strip().str.contains("STRTQ", case=False, na=False)]
        if strtq_candidates.empty:
            print(f"[!] STRTQ not found in sheet '{date_str}'. Skipping zone.")
            return

        strtq_index = strtq_candidates.index[0]
        strtq_value = df_sheet[strtq_index].iloc[1]

        new_column = []
        for row in data_lines:
            parts = row.strip().split()
            if len(parts) < 3:
                break
            try:
                val2 = float(parts[1])
                val3 = int(parts[2])
                if val3 == 1:
                    new_column.append(val2 + float(strtq_value))
            except:
                continue

        col_name = f"S{zone['scenario']}E{zone['event']}"
        df_out = data_frames[date_str]
        df_out[col_name] = pd.Series(new_column)

    # Loop through PSO lines
    for line in lines:
        line = line.strip()
        if line.startswith("ZONE T="):
            if current_zone:
                process_zone(current_zone, current_data)
            current_data = []
            current_zone = {}
            try:
                date_part = line.split('"')[1].split()[0]
                current_zone['date'] = date_part
                scenario_match = re.search(r"Scenario=\s*(\d+)", line)
                event_match = re.search(r"Event=\s*(\d+)", line)
                if scenario_match and event_match:
                    current_zone['scenario'] = scenario_match.group(1)
                    current_zone['event'] = event_match.group(1)
                else:
                    print(f"[!] Could not extract Scenario or Event in: {line}")
                    continue
            except:
                print(f"[!] Malformed ZONE line: {line}")
                continue
        elif current_zone:
            current_data.append(line)

    if current_zone:
        process_zone(current_zone, current_data)

    # Save output Excel
    for sheet_name, df in data_frames.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    writer.close()
    print(f"Done. Output saved to: {output_path}")
