import os
import pandas as pd
import numpy as np
import xlsxwriter

# Folder containing the Excel files
input_folder = '/home/pouria/git/water-institute/data/checking_similarity/input'
output_folder = '/home/pouria/git/water-institute/data/checking_similarity/opuput_files'
os.makedirs(output_folder, exist_ok=True)

# Process each Excel file
for filename in os.listdir(input_folder):
    if filename.endswith('.xlsx'):
        file_path = os.path.join(input_folder, filename)
        xls = pd.ExcelFile(file_path)

        # Output file
        output_path = os.path.join(output_folder, f'chart_{filename}')
        workbook = xlsxwriter.Workbook(output_path)

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)

            # Combine date and time as string
            df['datetime_str'] = df.iloc[:, 0].astype(str) + ' ' + df.iloc[:, 1].astype(str)
            df = df.reset_index(drop=True)

            # Create worksheet
            worksheet = workbook.add_worksheet(sheet_name[:31])
            header_format = workbook.add_format({'bold': True, 'font_name': 'Times New Roman'})

            # Write headers
            worksheet.write(0, 0, 'Datetime', header_format)
            for i, col in enumerate(df.columns[2:-1], start=1):  # Skip date, time, datetime_str
                worksheet.write(0, i, col, header_format)

            # Write data
            for row_idx, row in df.iterrows():
                worksheet.write(row_idx + 1, 0, row['datetime_str'])  # datetime string
                for col_idx, col in enumerate(df.columns[2:-1], start=1):
                    value = row[col]
                    if isinstance(value, (float, int)) and (np.isnan(value) or np.isinf(value)):
                        value = None
                    worksheet.write(row_idx + 1, col_idx, value)

            # Create a combined chart for all columns (except datetime)
            chart = workbook.add_chart({'type': 'line'})

            # Add all series to one chart
            for col_idx, col in enumerate(df.columns[2:-1], start=1):
                chart.add_series({
                    'name':       col,
                    'categories': [sheet_name[:31], 1, 0, len(df), 0],  # datetime_str
                    'values':     [sheet_name[:31], 1, col_idx, len(df), col_idx],
                    'line':       {'width': 2.25},
                })

            # Apply Times New Roman font to all chart text
            font = {'name': 'Times New Roman', 'size': 12}

            chart.set_title({
                'name': f"Event start date: {sheet_name}",
                'name_font': font
            })
            chart.set_x_axis({
                'name': 'Time (hour)',
                'name_font': font,
                'num_font': font,
            })
            chart.set_y_axis({
                'name': 'Discharge (cms)',
                'name_font': font,
                'num_font': font,
            })
            chart.set_legend({'font': font})
            chart.set_size({'width': 960, 'height': 480})

            # Insert chart to the right of the data table
            worksheet.insert_chart(0, len(df.columns) + 3, chart)

        workbook.close()

print("Charts created successfully with Times New Roman and larger size.")