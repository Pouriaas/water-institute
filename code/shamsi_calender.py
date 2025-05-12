# -*- coding: utf-8 -*-
"""
Created on Thu Jan 16 16:40:08 2025

@author: pouria
"""

import pandas as pd
import numpy as np
from persiantools.jdatetime import JalaliDate


def shamsi_hourly(start_date=None,end_date=None):

    # Define the start and end years in Solar Hijri
    start=start_date.split("/")
    start=[int(i) for i in start]
    start = JalaliDate(start[0], start[1], start[2]).toordinal()
    end=end_date.split("/")
    end=[int(i) for i in end]
    end = JalaliDate(end[0], end[1], end[2]).toordinal()
    
    
    # Initialize a list to store dates
    dates = []
    for i in range(start,end+1):
        date=str(JalaliDate.fromordinal(i)).replace("-", "/")
        dates.append(date)
    
    
    # Create a DataFrame with the generated dates
    calen = pd.DataFrame(dates, columns=["Solar Hijri Date"])
    rooz=np.full([len(calen)*24],np.nan,dtype=object)
    j=0
    for roozi in calen["Solar Hijri Date"]:
        for i in range(24):
            roozii=roozi[2:]+"/"+f"{i:02}:00"
            rooz[j]=roozii
            
            j+=1
    c=pd.DataFrame(rooz.T,columns=["shamsi_hourly"])
    return c

