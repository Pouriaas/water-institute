#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Jun 28 12:22:34 2025
Author: Pouria
"""

import numpy as np
import pandas as pd
import random
from sklearn.model_selection import LeaveOneOut
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error, r2_score
from scipy.stats import pearsonr
import statsmodels.api as sm
from collections import Counter, defaultdict

# -------------------------------
# Helper Functions
# -------------------------------
def adjusted_r2(model, X, y):
    r2 = model.rsquared
    n = X.shape[0]
    k = X.shape[1]
    return 1 - (1 - r2) * (n - 1) / (n - k - 1)

def stepwise_selection(X, y, method="adjusted_r2", 
                       threshold_in=0.01, 
                       threshold_out=0.05, 
                       initial_features=None,
                       max_features=7,
                       verbose=False):
    included = list(initial_features) if initial_features else []
    best_score = -np.inf if method == "adjusted_r2" else np.inf

    while True:
        changed = False
        if len(included) >= max_features:
            break

        excluded = list(set(X.columns) - set(included))
        scores_with_candidates = []

        for new_col in excluded:
            try:
                model = sm.OLS(y, sm.add_constant(X[included + [new_col]])).fit()

                if method == "pvalue":
                    pval = model.pvalues[new_col]
                    scores_with_candidates.append((pval, new_col))
                elif method == "adjusted_r2":
                    score = adjusted_r2(model, X[included + [new_col]], y)
                    scores_with_candidates.append((score, new_col))
                elif method == "AIC":
                    scores_with_candidates.append((model.aic, new_col))
                elif method == "BIC":
                    scores_with_candidates.append((model.bic, new_col))
            except:
                continue

        if not scores_with_candidates:
            break

        if method == "pvalue":
            scores_with_candidates.sort()
            best_new_score, best_candidate = scores_with_candidates[0]
            if best_new_score < threshold_in:
                included.append(best_candidate)
                changed = True
                if verbose:
                    print(f"Add {best_candidate:20s} | p-value = {best_new_score:.4f}")
        elif method == "adjusted_r2":
            scores_with_candidates.sort(reverse=True)
            best_new_score, best_candidate = scores_with_candidates[0]
            if best_new_score > best_score:
                best_score = best_new_score
                included.append(best_candidate)
                changed = True
                if verbose:
                    print(f"Add {best_candidate:20s} | adj R² = {best_new_score:.4f}")
        elif method in ["AIC", "BIC"]:
            scores_with_candidates.sort()
            best_new_score, best_candidate = scores_with_candidates[0]
            if best_new_score < best_score:
                best_score = best_new_score
                included.append(best_candidate)
                changed = True
                if verbose:
                    print(f"Add {best_candidate:20s} | {method} = {best_new_score:.4f}")

        if not changed:
            break

    return included

# -------------------------------
# Remove Highly Correlated Features
# -------------------------------
def remove_highly_correlated_features(X, threshold=0.95):
    corr_matrix = X.corr().abs()
    upper = corr_matrix.where(np.triu(np.ones(corr_matrix.shape), k=1).astype(bool))
    to_drop = [column for column in upper.columns if any(upper[column] >= threshold)]
    return X.drop(columns=to_drop), to_drop

# -------------------------------
# LOOCV Stepwise Evaluation
# -------------------------------
def loocv_stepwise(X, y, method="adjusted_r2", initial_features=None, max_features=7):
    loo = LeaveOneOut()
    y_true, y_pred = [], []
    selected_counts = Counter()
    coef_sum = defaultdict(float)

    for train_index, test_index in loo.split(X):
        X_train, X_test = X.iloc[train_index], X.iloc[test_index]
        y_train, y_test = y.iloc[train_index], y.iloc[test_index]

        selected_features = stepwise_selection(X_train, y_train, 
                                               method=method,
                                               initial_features=initial_features,
                                               max_features=max_features,
                                               verbose=False)

        for feat in selected_features:
            selected_counts[feat] += 1

        model = LinearRegression()
        model.fit(X_train[selected_features], y_train)

        for i, feat in enumerate(selected_features):
            coef_sum[feat] += abs(model.coef_[i])

        y_hat = model.predict(X_test[selected_features])
        y_pred.append(y_hat[0])
        y_true.append(y_test.iloc[0])

    mse = mean_squared_error(y_true, y_pred)
    r2 = r2_score(y_true, y_pred)
    y_true_np, y_pred_np = np.array(y_true), np.array(y_pred)
    y_mean = np.mean(y_true_np)
    nse = 1 - np.sum((y_true_np - y_pred_np) ** 2) / np.sum((y_true_np - y_mean) ** 2)
    r, _ = pearsonr(y_true_np, y_pred_np)
    beta = np.mean(y_pred_np) / np.mean(y_true_np)
    gamma = np.std(y_pred_np) / np.std(y_true_np)
    kge = 1 - np.sqrt((r - 1) ** 2 + (beta - 1) ** 2 + (gamma - 1) ** 2)

    avg_coefs = {feat: coef_sum[feat] / selected_counts[feat] for feat in selected_counts}

    return mse, r2, nse, kge, y_true, y_pred, selected_counts, avg_coefs

# -------------------------------
# Main Execution
# -------------------------------
if __name__ == "__main__":
    df = pd.read_excel("/mnt/c/Users/pouria/Desktop/UTWI/basins_charactristics/regenalization/shomal.xlsx", sheet_name="Sheet3")
    y = df["Tr"]
    X = df.drop(columns=["Tr", "Point_ID"])

    # Remove correlated features
    X, dropped = remove_highly_correlated_features(X, threshold=0.95)

    # Suggested features based on your correlation table
    suggested = ["Kirpich", "Slop (-)", "Farm_Land", "Uncovered_Plain", "Water", "Urban",]

    # LOOCV Stepwise
    mse, r2, nse, kge, y_true, y_hat, selected_counts, avg_coefs = loocv_stepwise(
        X, y, method="adjusted_r2", initial_features=suggested, max_features=7
    )

    # Print results
    print("\nModel Evaluation Metrics:")
    print(f"  LOOCV MSE : {mse:.4f}")
    print(f"  R²        : {r2:.4f}")
    print(f"  NSE       : {nse:.4f}")
    print(f"  KGE       : {kge:.4f}")

    print(f"\nTotal unique features selected across LOOCV folds: {len(selected_counts)}")
    print("Feature selection frequency and average |coefficients|:")
    for feat, count in selected_counts.most_common():
        print(f"  {feat:20s} selected in {count:2d} folds, avg |coef| = {avg_coefs[feat]:.4f}")