#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Jun 29 09:14:04 2025

@author: pouria
"""

import pandas as pd
import numpy as np
from sklearn.model_selection import LeaveOneOut
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error, r2_score
from scipy.stats import pearsonr
from sklearn.neighbors import KNeighborsRegressor

# Load your data
df = pd.read_excel("/mnt/c/Users/pouria/Desktop/UTWI/basins_charactristics/regenalization/shomal+namak.xlsx", sheet_name="Sheet2")

# Define target and selected features
target = "Tc"
selected_features = ['Urban', 'silt 0-5cm (%) p95', 'Kirpich']  # ✅ You can adjust this list

# Drop rows with NaN in selected columns
df_clean = df[[target] + selected_features].dropna()

# Prepare data
X = df_clean[selected_features]
y = df_clean[target]

# Initialize LOOCV
loo = LeaveOneOut()
y_true, y_pred = [], []

# Loop over LOOCV
# Inside the LOOCV loop
for train_index, test_index in loo.split(X):
    X_train, X_test = X.iloc[train_index], X.iloc[test_index]
    y_train, y_test = y[train_index], y[test_index]

    model = LinearRegression()
    model.fit(X_train, y_train)
    y_hat = model.predict(X_test)

    y_true.append(y_test.iloc[0])  # ✅ Use iloc here
    y_pred.append(y_hat[0])

# Convert to arrays
y_true = np.array(y_true)
y_pred = np.array(y_pred)

Tc=[]

Tc.append(df["Point_ID"])
Tc.append(y_true)
Tc.append(y_pred)

TC=pd.DataFrame(np.array(Tc).T,columns=["Point_ID","True","Sim"])


# Load your data
df = pd.read_excel("/mnt/c/Users/pouria/Desktop/UTWI/basins_charactristics/regenalization/shomal+namak.xlsx", sheet_name="Sheet3")

# Define target and selected features
target = "Tr"
selected_features = ['soiltype_p95', 'GCNp50', 'Kirpich']  # ✅ You can adjust this list

# Drop rows with NaN in selected columns
df_clean = df[[target] + selected_features].dropna()

# Prepare data
X = df_clean[selected_features]
y = df_clean[target]

# Initialize LOOCV
loo = LeaveOneOut()
y_true, y_pred = [], []

# Loop over LOOCV
# Inside the LOOCV loop
for train_index, test_index in loo.split(X):
    X_train, X_test = X.iloc[train_index], X.iloc[test_index]
    y_train, y_test = y[train_index], y[test_index]

    model = KNeighborsRegressor(n_neighbors=3, weights="uniform", p=2)
    model.fit(X_train, y_train)
    y_hat = model.predict(X_test)

    y_true.append(y_test.iloc[0])  # ✅ Use iloc here
    y_pred.append(y_hat[0])

# Convert to arrays
y_true = np.array(y_true)
y_pred = np.array(y_pred)

Tr=[]

Tr.append(df["Point_ID"])
Tr.append(y_true)
Tr.append(y_pred)

###
#('soiltype_p95', 'GCNp50', 'Kirpich')	3	distance	2
selected_features = ['soiltype_p95', 'GCNp50', 'Kirpich']
# Drop rows with NaN in selected columns
df_clean = df[[target] + selected_features].dropna()

# Prepare data
X = df_clean[selected_features]
y = df_clean[target]

# Initialize LOOCV
loo = LeaveOneOut()
y_true, y_pred = [], []

# Loop over LOOCV
# Inside the LOOCV loop
for train_index, test_index in loo.split(X):
    X_train, X_test = X.iloc[train_index], X.iloc[test_index]
    y_train, y_test = y[train_index], y[test_index]

    model = KNeighborsRegressor(n_neighbors=3, weights="distance", p=2)
    model.fit(X_train, y_train)
    y_hat = model.predict(X_test)

    y_pred.append(y_hat[0])

# Convert to arrays
y_pred = np.array(y_pred)


Tr.append(y_pred)
###
#('lon', 'soiltype_p95', 'Urban')	3	uniform	1	
###
selected_features = ['lon', 'soiltype_p95', 'Urban']
# Drop rows with NaN in selected columns
df_clean = df[[target] + selected_features].dropna()

# Prepare data
X = df_clean[selected_features]
y = df_clean[target]

# Initialize LOOCV
loo = LeaveOneOut()
y_true, y_pred = [], []

# Loop over LOOCV
# Inside the LOOCV loop
for train_index, test_index in loo.split(X):
    X_train, X_test = X.iloc[train_index], X.iloc[test_index]
    y_train, y_test = y[train_index], y[test_index]

    model = KNeighborsRegressor(n_neighbors=3, weights="uniform", p=1)
    model.fit(X_train, y_train)
    y_hat = model.predict(X_test)

    y_pred.append(y_hat[0])

# Convert to arrays
y_pred = np.array(y_pred)


Tr.append(y_pred)

TR=pd.DataFrame(np.array(Tr).T,columns=["Point_ID","True","Sim1","Sim2","Sim3"])



# Load your data
df = pd.read_excel("/mnt/c/Users/pouria/Desktop/UTWI/basins_charactristics/regenalization/shomal.xlsx", sheet_name="Sheet4")

# Define target and selected features
target = "STRKR"
selected_features = ['clay 0-5cm (%) mean', 'clay 0-5cm (%) p5', 'silt 15-30cm (%) p95']  # ✅ You can adjust this list
df_clean = df[[target] + selected_features].dropna()

# Prepare data
X = df_clean[selected_features]
y = df_clean[target]

# Initialize LOOCV
loo = LeaveOneOut()
y_true, y_pred = [], []

# Loop over LOOCV
# Inside the LOOCV loop
for train_index, test_index in loo.split(X):
    X_train, X_test = X.iloc[train_index], X.iloc[test_index]
    y_train, y_test = y[train_index], y[test_index]

    model = KNeighborsRegressor(n_neighbors=3, weights="uniform", p=2)
    model.fit(X_train, y_train)
    y_hat = model.predict(X_test)

    y_true.append(y_test.iloc[0])  # ✅ Use iloc here
    y_pred.append(y_hat[0])

# Convert to arrays
y_true = np.array(y_true)
y_pred = np.array(y_pred)

STRKR=[]

STRKR.append(df["Point_ID"])
STRKR.append(y_true)
STRKR.append(y_pred)
####
# 'GCNp50', 'sand 0-5cm (%) p5', 'silt 100-200cm (%) p95'
# 3	uniform	2
####
selected_features = ['GCNp50', 'sand 0-5cm (%) p5', 'silt 100-200cm (%) p95']  # ✅ You can adjust this list
df_clean = df[[target] + selected_features].dropna()

# Prepare data
X = df_clean[selected_features]
y = df_clean[target]

# Initialize LOOCV
loo = LeaveOneOut()
y_true, y_pred = [], []

# Loop over LOOCV
# Inside the LOOCV loop
for train_index, test_index in loo.split(X):
    X_train, X_test = X.iloc[train_index], X.iloc[test_index]
    y_train, y_test = y[train_index], y[test_index]

    model = KNeighborsRegressor(n_neighbors=3, weights="uniform", p=2)
    model.fit(X_train, y_train)
    y_hat = model.predict(X_test)

    y_pred.append(y_hat[0])

# Convert to arrays
y_pred = np.array(y_pred)
STRKR.append(y_pred)

# ('Outcrop', 'silt 15-30cm (%) p95', 'silt 30-60cm (%) p95')	3	distance	2
#####
# Drop rows with NaN in selected columns
selected_features = ['Outcrop', 'silt 15-30cm (%) p95', 'silt 30-60cm (%) p95']  # ✅ You can adjust this list

df_clean = df[[target] + selected_features].dropna()

# Prepare data
X = df_clean[selected_features]
y = df_clean[target]

# Initialize LOOCV
loo = LeaveOneOut()
y_true, y_pred = [], []

# Loop over LOOCV
# Inside the LOOCV loop
for train_index, test_index in loo.split(X):
    X_train, X_test = X.iloc[train_index], X.iloc[test_index]
    y_train, y_test = y[train_index], y[test_index]

    model = KNeighborsRegressor(n_neighbors=3, weights="distance", p=2)
    model.fit(X_train, y_train)
    y_hat = model.predict(X_test)

    y_pred.append(y_hat[0])

# Convert to arrays
y_pred = np.array(y_pred)
STRKR.append(y_pred)

STRKRR=pd.DataFrame(np.array(STRKR).T,columns=["Point_ID","True","Sim1","Sim2","Sim3"])





# Load your data
df = pd.read_excel("/mnt/c/Users/pouria/Desktop/UTWI/basins_charactristics/regenalization/shomal.xlsx", sheet_name="Sheet5")

# Define target and selected features
target = "DLTKR"


# ('lat', 'Marshland', 'sand 0-5cm (%) mean')	1	distance	1

selected_features = ['lat', 'Marshland', 'sand 0-5cm (%) mean']  # ✅ You can adjust this list
# Drop rows with NaN in selected columns
df_clean = df[[target] + selected_features].dropna()

# Prepare data
X = df_clean[selected_features]
y = df_clean[target]

# Initialize LOOCV
loo = LeaveOneOut()
y_true, y_pred = [], []

# Loop over LOOCV
# Inside the LOOCV loop
for train_index, test_index in loo.split(X):
    X_train, X_test = X.iloc[train_index], X.iloc[test_index]
    y_train, y_test = y[train_index], y[test_index]

    model = KNeighborsRegressor(n_neighbors=1, weights="distance", p=1)
    model.fit(X_train, y_train)
    y_hat = model.predict(X_test)

    y_true.append(y_test.iloc[0])  # ✅ Use iloc here
    y_pred.append(y_hat[0])

# Convert to arrays
y_true = np.array(y_true)
y_pred = np.array(y_pred)

DLTKR=[]

DLTKR.append(df["Point_ID"])
DLTKR.append(y_true)
DLTKR.append(y_pred)

# ('lat', 'soiltype_p95', 'Marshland')	3	distance	2

selected_features = ['lat', 'soiltype_p95', 'Marshland']  # ✅ You can adjust this list
# Drop rows with NaN in selected columns
df_clean = df[[target] + selected_features].dropna()

# Prepare data
X = df_clean[selected_features]
y = df_clean[target]

# Initialize LOOCV
loo = LeaveOneOut()
y_true, y_pred = [], []

# Loop over LOOCV
# Inside the LOOCV loop
for train_index, test_index in loo.split(X):
    X_train, X_test = X.iloc[train_index], X.iloc[test_index]
    y_train, y_test = y[train_index], y[test_index]

    model = KNeighborsRegressor(n_neighbors=3, weights="distance", p=2)
    model.fit(X_train, y_train)
    y_hat = model.predict(X_test)

    y_true.append(y_test.iloc[0])  # ✅ Use iloc here
    y_pred.append(y_hat[0])

# Convert to arrays
y_true = np.array(y_true)
y_pred = np.array(y_pred)

DLTKR.append(y_pred)

# ('lat', 'soiltype_p95', 'Marshland')	3	uniform	2

selected_features = ['lat', 'soiltype_p95', 'Marshland']  # ✅ You can adjust this list
# Drop rows with NaN in selected columns
df_clean = df[[target] + selected_features].dropna()

# Prepare data
X = df_clean[selected_features]
y = df_clean[target]

# Initialize LOOCV
loo = LeaveOneOut()
y_true, y_pred = [], []

# Loop over LOOCV
# Inside the LOOCV loop
for train_index, test_index in loo.split(X):
    X_train, X_test = X.iloc[train_index], X.iloc[test_index]
    y_train, y_test = y[train_index], y[test_index]

    model = KNeighborsRegressor(n_neighbors=3, weights="uniform", p=2)
    model.fit(X_train, y_train)
    y_hat = model.predict(X_test)

    y_true.append(y_test.iloc[0])  # ✅ Use iloc here
    y_pred.append(y_hat[0])

# Convert to arrays
y_true = np.array(y_true)
y_pred = np.array(y_pred)

DLTKR.append(y_pred)
DLTKRR=pd.DataFrame(np.array(DLTKR).T,columns=["Point_ID","True","Sim1","Sim2","Sim3"])


# Load your data
df = pd.read_excel("/mnt/c/Users/pouria/Desktop/UTWI/basins_charactristics/regenalization/shomal+namak.xlsx", sheet_name="Sheet6")

# Define target and selected features
target = "RTIOL"

# ('lon', 'Urban', 'Marshland')	3	distance	2

selected_features = ['lon', 'Urban', 'Marshland']  # ✅ You can adjust this list
# Drop rows with NaN in selected columns
df_clean = df[[target] + selected_features].dropna()

# Prepare data
X = df_clean[selected_features]
y = df_clean[target]

# Initialize LOOCV
loo = LeaveOneOut()
y_true, y_pred = [], []

# Loop over LOOCV
# Inside the LOOCV loop
for train_index, test_index in loo.split(X):
    X_train, X_test = X.iloc[train_index], X.iloc[test_index]
    y_train, y_test = y[train_index], y[test_index]

    model = KNeighborsRegressor(n_neighbors=3, weights="distance", p=2)
    model.fit(X_train, y_train)
    y_hat = model.predict(X_test)

    y_true.append(y_test.iloc[0])  # ✅ Use iloc here
    y_pred.append(y_hat[0])

# Convert to arrays
y_true = np.array(y_true)
y_pred = np.array(y_pred)

RTIOL=[]

RTIOL.append(df["Point_ID"])
RTIOL.append(y_true)
RTIOL.append(y_pred)

# ('lon', 'Urban', 'Marshland')	3	uniform	2

selected_features = ['lon', 'Urban', 'Marshland']  # ✅ You can adjust this list
# Drop rows with NaN in selected columns
df_clean = df[[target] + selected_features].dropna()

# Prepare data
X = df_clean[selected_features]
y = df_clean[target]

# Initialize LOOCV
loo = LeaveOneOut()
y_true, y_pred = [], []

# Loop over LOOCV
# Inside the LOOCV loop
for train_index, test_index in loo.split(X):
    X_train, X_test = X.iloc[train_index], X.iloc[test_index]
    y_train, y_test = y[train_index], y[test_index]

    model = KNeighborsRegressor(n_neighbors=3, weights="uniform", p=2)
    model.fit(X_train, y_train)
    y_hat = model.predict(X_test)

    y_true.append(y_test.iloc[0])  # ✅ Use iloc here
    y_pred.append(y_hat[0])

# Convert to arrays
y_true = np.array(y_true)
y_pred = np.array(y_pred)

RTIOL.append(y_pred)

# ('lon', 'Marshland', 'Uncovered_Plain')	3	distance	2


selected_features = ['lon', 'Marshland', 'Uncovered_Plain']  # ✅ You can adjust this list
# Drop rows with NaN in selected columns
df_clean = df[[target] + selected_features].dropna()

# Prepare data
X = df_clean[selected_features]
y = df_clean[target]

# Initialize LOOCV
loo = LeaveOneOut()
y_true, y_pred = [], []

# Loop over LOOCV
# Inside the LOOCV loop
for train_index, test_index in loo.split(X):
    X_train, X_test = X.iloc[train_index], X.iloc[test_index]
    y_train, y_test = y[train_index], y[test_index]

    model = KNeighborsRegressor(n_neighbors=3, weights="distance", p=2)
    model.fit(X_train, y_train)
    y_hat = model.predict(X_test)

    y_true.append(y_test.iloc[0])  # ✅ Use iloc here
    y_pred.append(y_hat[0])

# Convert to arrays
y_true = np.array(y_true)
y_pred = np.array(y_pred)

RTIOL.append(y_pred)
RTIOLR=pd.DataFrame(np.array(RTIOL).T,columns=["Point_ID","True","Sim1","Sim2","Sim3"])


# Load your data
df = pd.read_excel("/mnt/c/Users/pouria/Desktop/UTWI/basins_charactristics/regenalization/shomal.xlsx", sheet_name="Sheet12")

# Define target and selected features
target = "ERAIN"

# ('lat', 'silt 15-30cm (%) p5', 'Slop (-)')	3	distance	1

selected_features = ['lat', 'silt 15-30cm (%) p5', 'Slop (-)']  # ✅ You can adjust this list
# Drop rows with NaN in selected columns
df_clean = df[[target] + selected_features].dropna()

# Prepare data
X = df_clean[selected_features]
y = df_clean[target]

# Initialize LOOCV
loo = LeaveOneOut()
y_true, y_pred = [], []

# Loop over LOOCV
# Inside the LOOCV loop
for train_index, test_index in loo.split(X):
    X_train, X_test = X.iloc[train_index], X.iloc[test_index]
    y_train, y_test = y[train_index], y[test_index]

    model = KNeighborsRegressor(n_neighbors=3, weights="distance", p=1)
    model.fit(X_train, y_train)
    y_hat = model.predict(X_test)

    y_true.append(y_test.iloc[0])  # ✅ Use iloc here
    y_pred.append(y_hat[0])

# Convert to arrays
y_true = np.array(y_true)
y_pred = np.array(y_pred)

ERAIN=[]

ERAIN.append(df["Point_ID"])
ERAIN.append(y_true)
ERAIN.append(y_pred)

# ('lat', 'Water', 'silt 15-30cm (%) p5')	3	distance	1

selected_features = ['lat', 'Water', 'silt 15-30cm (%) p5']  # ✅ You can adjust this list
# Drop rows with NaN in selected columns
df_clean = df[[target] + selected_features].dropna()

# Prepare data
X = df_clean[selected_features]
y = df_clean[target]

# Initialize LOOCV
loo = LeaveOneOut()
y_true, y_pred = [], []

# Loop over LOOCV
# Inside the LOOCV loop
for train_index, test_index in loo.split(X):
    X_train, X_test = X.iloc[train_index], X.iloc[test_index]
    y_train, y_test = y[train_index], y[test_index]

    model = KNeighborsRegressor(n_neighbors=3, weights="distance", p=1)
    model.fit(X_train, y_train)
    y_hat = model.predict(X_test)

    y_true.append(y_test.iloc[0])  # ✅ Use iloc here
    y_pred.append(y_hat[0])

# Convert to arrays
y_true = np.array(y_true)
y_pred = np.array(y_pred)

ERAIN.append(y_pred)

# ('lat', 'silt 15-30cm (%) p5', 'Slop (-)')	3	uniform	1


selected_features = ['lat', 'silt 15-30cm (%) p5', 'Slop (-)']  # ✅ You can adjust this list
# Drop rows with NaN in selected columns
df_clean = df[[target] + selected_features].dropna()

# Prepare data
X = df_clean[selected_features]
y = df_clean[target]

# Initialize LOOCV
loo = LeaveOneOut()
y_true, y_pred = [], []

# Loop over LOOCV
# Inside the LOOCV loop
for train_index, test_index in loo.split(X):
    X_train, X_test = X.iloc[train_index], X.iloc[test_index]
    y_train, y_test = y[train_index], y[test_index]

    model = KNeighborsRegressor(n_neighbors=3, weights="uniform", p=1)
    model.fit(X_train, y_train)
    y_hat = model.predict(X_test)

    y_true.append(y_test.iloc[0])  # ✅ Use iloc here
    y_pred.append(y_hat[0])

# Convert to arrays
y_true = np.array(y_true)
y_pred = np.array(y_pred)

ERAIN.append(y_pred)
ERAINR=pd.DataFrame(np.array(ERAIN).T,columns=["Point_ID","True","Sim1","Sim2","Sim3"])

# Dictionary of sheet names and DataFrames
dfs = {
    'Tc': TC,
    'Tr': TR,
    'STRKR': STRKRR,
    'DLTKR': DLTKRR,
    'RTIOL': RTIOLR,
    'ERAIN': ERAINR
}

# Write to Excel file
output_file = '/mnt/c/Users/pouria/Desktop/UTWI/basins_charactristics/regenalization/selected_var/multiple_sheets.xlsx'
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for sheet_name, dataframe in dfs.items():
        dataframe.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Excel file '{output_file}' created with {len(dfs)} sheets.")




