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

# Load your data
df = pd.read_excel("/mnt/c/Users/pouria/Desktop/UTWI/basins_charactristics/regenalization/shomal.xlsx", sheet_name="Sheet3")

# Define target and selected features
target = "Tr"
selected_features = ["Kirpich", "Water", "Farm_Land"]  # ✅ You can adjust this list

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

# Evaluation Metrics
mse = mean_squared_error(y_true, y_pred)
r2 = r2_score(y_true, y_pred)
nse = 1 - np.sum((y_true - y_pred) ** 2) / np.sum((y_true - np.mean(y_true)) ** 2)
r, _ = pearsonr(y_true, y_pred)
beta = np.mean(y_pred) / np.mean(y_true)
gamma = np.std(y_pred) / np.std(y_true)
kge = 1 - np.sqrt((r - 1) ** 2 + (beta - 1) ** 2 + (gamma - 1) ** 2)

# Output
print("Linear Regression with 4 Features (LOOCV):")
print(f"  MSE  = {mse:.4f}")
print(f"  R²   = {r2:.4f}")
print(f"  NSE  = {nse:.4f}")
print(f"  KGE  = {kge:.4f}")