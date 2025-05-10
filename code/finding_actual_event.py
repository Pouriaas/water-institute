import numpy as np
import pandas as pd
import os 
import openpyxl

input_file=r'/home/pouria/git/water-institute/data'

stations_name=["18-251_hourly.xlsx","18-061_hourly.xlsx","18-003_hourly.xlsx"]

shamsi_calen=pd.read_excel(r"/home/pouria/git/water-institute/data/solar_hijri_dates_1350_to_1401.xlsx")
array=np.full((len (stations_name)+1,len(shamsi_calen)),np.nan,dtype="object")

df=pd.DataFrame(array.T,columns=["Date"] + stations_name)
df["Date"]=shamsi_calen
df=df.set_index("Date")

for file_name in stations_name:
    path=os.path.join(input_file,file_name)
    wb = openpyxl.load_workbook(path)

    # Define the target color (purple) in Excel ARGB format
    # Example: purple RGB (128, 0, 128) → hex '800080' → ARGB 'FF800080'
    target_color = 'FF7030A0'

    # List to store matching sheet names
    purple_sheets = []

    # Loop over all sheets
    for sheetname in wb.sheetnames:
        sheet = wb[sheetname]
        tab_color = sheet.sheet_properties.tabColor

        if tab_color is not None:
            tab_color_rgb = tab_color.rgb
            if tab_color_rgb and tab_color_rgb.upper() == target_color:
                purple_sheets.append(sheetname)
    for sheet_name in purple_sheets:
        sim=pd.read_excel(path,sheet_name=sheet_name)
        valid_date=sim["Date"].dropna().unique()
        sim=sim.set_index("Date")
        for date in valid_date:
            if pd.isna(sim["Time"][date]):
                valid_date=valid_date.drop(date)
            else:
                year=date.split("/")[0]
                if str(year)=="00":
                    year="1400"
                else:
                    year="13"+str(year)
                newformat=date.split("/")
                new_date=year+"/"+newformat[1]+"/"+newformat[2]
                df.loc[new_date, file_name] = 1
df = df.dropna(how='all')

    