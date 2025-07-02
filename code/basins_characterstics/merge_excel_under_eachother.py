#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jun 24 14:28:21 2025

@author: pouria
"""

import os
import pandas as pd



folders=os.listdir(r"/home/pouria/git/water-institute/data/basins_charactristics/output/excel/41")
files=os.listdir(r"/home/pouria/git/water-institute/data/basins_charactristics/output/excel/41/411")


output=r"/home/pouria/git/water-institute/data/basins_charactristics/output/excel/41"

for file in files:
    all_dfs = []
    for folder in folders:
        fifi=os.path.join(output,folder)
        df = pd.read_excel(os.path.join(fifi, file))
        all_dfs.append(df)
    merged_df = pd.concat(all_dfs, ignore_index=True)
    
    merged_df.to_excel(os.path.join(output, file), index=False)
        
    
