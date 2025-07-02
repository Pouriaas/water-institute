#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Jun 29 10:47:10 2025
@author: pouria
"""

import pandas as pd
import numpy as np
from sklearn.model_selection import LeaveOneOut
from sklearn.neighbors import KNeighborsRegressor
from sklearn.metrics import r2_score, mean_squared_error
from scipy.stats import pearsonr
from itertools import combinations
from joblib import Parallel, delayed
import multiprocessing
from itertools import product  # to loop over param grid

# Custom metrics
def nse_score(y_true, y_pred):
    return 1 - np.sum((y_true - y_pred) ** 2) / np.sum((y_true - np.mean(y_true)) ** 2)

def kge_score(y_true, y_pred):
    r, _ = pearsonr(y_true, y_pred)
    beta = np.mean(y_pred) / (np.mean(y_true) + 1e-10)
    gamma = np.std(y_pred) / (np.std(y_true) + 1e-10)
    return 1 - np.sqrt((r - 1)**2 + (beta - 1)**2 + (gamma - 1)**2)
# Define your hyperparameter grid
param_grid = {
    'n_neighbors': [1, 3, 5, 7, 9],
    'weights': ['uniform', 'distance'],
    'p': [1, 2]
}

# Load data
df = pd.read_excel("/mnt/c/Users/pouria/Desktop/UTWI/basins_charactristics/regenalization/shomal+namak.xlsx", sheet_name="Sheet2")
X_raw = df.drop(columns=["Tc", "Point_ID"], errors='ignore')
y = df["Tc"]

# Drop constant columns
X_raw = X_raw.loc[:, X_raw.nunique() > 1]

# Step 1: Correlation filter (|r| > 0.2)
corr_with_target = X_raw.corrwith(y).dropna()
selected_by_corr = corr_with_target[abs(corr_with_target) > 0.20].index.tolist()
X_filtered = X_raw[selected_by_corr]

# Step 2: Remove highly correlated features (|r| > 0.95)
corr_matrix = X_filtered.corr().abs()
upper_tri = corr_matrix.where(np.triu(np.ones(corr_matrix.shape), k=1).astype(bool))
to_drop = [col for col in upper_tri.columns if any(upper_tri[col] > 0.95)]
X_final = X_filtered.drop(columns=to_drop)

# Evaluation function using KNN
def evaluate_group_knn_with_params(group, X_final, y, param_grid):
    X_sub = X_final[list(group)]
    results = []

    # Loop over all hyperparameter combinations
    for n_neighbors, weights, p in product(param_grid['n_neighbors'], param_grid['weights'], param_grid['p']):
        loo = LeaveOneOut()
        y_true_loo, y_pred_loo = [], []

        for train_idx, test_idx in loo.split(X_sub):
            X_train, X_test = X_sub.iloc[train_idx], X_sub.iloc[test_idx]
            y_train, y_test = y.iloc[train_idx], y.iloc[test_idx]

            model = KNeighborsRegressor(n_neighbors=n_neighbors, weights=weights, p=p)
            model.fit(X_train, y_train)
            pred = model.predict(X_test)

            y_true_loo.append(y_test.values[0])
            y_pred_loo.append(pred[0])

        y_true_loo = np.array(y_true_loo)
        y_pred_loo = np.array(y_pred_loo)
        mse_loo = mean_squared_error(y_true_loo, y_pred_loo)
        r2_loo = r2_score(y_true_loo, y_pred_loo)
        nse_loo = nse_score(y_true_loo, y_pred_loo)
        kge_loo = kge_score(y_true_loo, y_pred_loo)

        # Full training performance
        model = KNeighborsRegressor(n_neighbors=n_neighbors, weights=weights, p=p)
        model.fit(X_sub, y)
        y_pred_train = model.predict(X_sub)
        mse_train = mean_squared_error(y, y_pred_train)
        r2_train = r2_score(y, y_pred_train)
        nse_train = nse_score(y, y_pred_train)
        kge_train = kge_score(y, y_pred_train)

        results.append({
            "features": group,
            "n_neighbors": n_neighbors,
            "weights": weights,
            "p": p,
            "mse_loo": mse_loo,
            "r2_loo": r2_loo,
            "nse_loo": nse_loo,
            "kge_loo": kge_loo,
            "mse_train": mse_train,
            "r2_train": r2_train,
            "nse_train": nse_train,
            "kge_train": kge_train,
        })

    return results

# Generate all 4-feature combinations
combs = list(combinations(X_final.columns, 3))
print(f"Total combinations: {len(combs)}")

# Parallel evaluation
n_jobs = max(1, multiprocessing.cpu_count() - 1)
all_results_nested = Parallel(n_jobs=n_jobs, verbose=5)(
    delayed(evaluate_group_knn_with_params)(group, X_final, y, param_grid) for group in combs
)

# Flatten results
all_results = [item for sublist in all_results_nested for item in sublist]

# Save
results_df = pd.DataFrame(all_results)
results_df_sorted = results_df.sort_values(by="r2_loo", ascending=False)
# Save results
results_df_sorted.to_excel("/mnt/c/Users/pouria/Desktop/UTWI/basins_charactristics/regenalization/shomal_namk/top_4_feature_models_parallel_knn_Tc.xlsx", index=False)
