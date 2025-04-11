# -*- coding: utf-8 -*-
"""
Created on Fri Apr 11 11:28:43 2025

@author: Gail McConnell, SIPBS, University of Strathclyde, UK
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.signal import argrelextrema

# Load the Excel file
def load_excel(file_path):
    df = pd.read_excel(file_path)
    return df

# Compute the mean range statistics using local minima and maxima
def compute_range_statistics(df):
    col_names = df.columns[1:3]

    range_stats = {}
    std_range_stats = {}
    peak_counts = {}

    for col in col_names:
        data = df[col].values

        # Find local minima and maxima
        local_max_indices = argrelextrema(data, np.greater)[0]
        local_min_indices = argrelextrema(data, np.less)[0]

        if len(local_max_indices) == 0 or len(local_min_indices) == 0:
            print(f"Warning: No local extrema found for column {col}.")
            continue

        local_maxima = data[local_max_indices]
        local_minima = data[local_min_indices]

        # Ensure same length for range calculation
        min_length = min(len(local_maxima), len(local_minima))
        local_ranges = np.abs(local_maxima[:min_length] - local_minima[:min_length])

        mean_range = np.mean(local_ranges)
        std_range = np.std(local_ranges, ddof=1)

        range_stats[col] = mean_range
        std_range_stats[col] = std_range
        peak_counts[col] = len(local_max_indices)  

    return range_stats, std_range_stats, peak_counts

# Plot the data
def plot_data(df):
    plt.figure(figsize=(10, 5))

    x = df.iloc[:, 0]
    y1 = df.iloc[:, 1]
    y2 = df.iloc[:, 2]

    plt.plot(x, y1, label=df.columns[1], linestyle='-', marker='o', alpha=0.7)
    plt.plot(x, y2, label=df.columns[2], linestyle='-', marker='s', alpha=0.7)

    plt.xlabel(df.columns[0])
    plt.ylabel("Normalised intensity (arb. units)")
    plt.legend()
    plt.title("Data Plot")
    plt.grid()
    plt.show()

# Calculate contrast improvement percentage
def calculate_contrast_improvement(val1, val2):
    lower, higher = sorted([val1, val2])
    improvement = ((higher - lower) / lower) * 100
    return improvement

# Main function
def main():
    file_path = r'C:\Spyder\Fishscale.xlsx'  # Update path and filename
    df = load_excel(file_path)
    range_stats, std_ranges, peak_counts = compute_range_statistics(df)

    print("Mean Range Statistics:")
    for col, r in range_stats.items():
        print(f"{col}: {r:.4f}")

    print("\nStandard Deviation of Ranges:")
    for col, s in std_ranges.items():
        print(f"{col}: {s:.4f}")

    print("\nNumber of Peaks Detected:")
    for col, p in peak_counts.items():
        print(f"{col}: {p}")

    # Contrast improvement (based on mean range)
    mean_values = list(range_stats.values())
    col_labels = list(range_stats.keys())

    if len(mean_values) == 2:
        improvement = calculate_contrast_improvement(mean_values[0], mean_values[1])
        better_label = col_labels[0] if mean_values[0] > mean_values[1] else col_labels[1]
        print(f"\nContrast in '{better_label}' is {improvement:.2f}% higher than the other dataset.")

    plot_data(df)

if __name__ == "__main__":
    main()
