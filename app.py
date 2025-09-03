import streamlit as st
import pandas as pd

# --- Load Excel Config ---
excel_file = "LNE CustomerHealthScoringModel.xlsx"

# Load Input sheet raw (no assumption about headers yet)
raw_df = pd.read_excel(excel_file, sheet_name="Input", header=None)

# Find the row that contains the true headers (look for 'Low Risk')
header_row = raw_df[raw_df.apply(lambda r: r.astype(str).str.contains("Low Risk").any(), axis=1)].index[0]

# Re-load with proper headers
df = pd.read_excel(excel_file, sheet_name="Input", header=header_row)

# --- Clean up columns ---
# Drop duplicate columns
df = df.loc[:, ~df.columns.duplicated()].copy()

# Keep only important ones
keep_cols = ["Metric", "Low Risk", "Moderate Risk", "High Risk", "Weight"]
if "Section" in df.columns:
    keep_cols.append("Section")

df = df[keep_cols]

# Convert thresholds and weights to numeric
numeric_cols = [c for c in ["Low Risk", "Moderate Risk", "High Risk", "Weight"] if c in df.columns]
df[numeric_cols] = df[numeric_cols].apply(pd.to_numeric, errors="coerce")

# Ensure Section column exists and is forward-filled
if "Section" not in df.columns:
    df["Section"] = df["Metric"].where(df["Low Risk"].isna()).ffill()

# Keep only rows with valid thresholds + weights
kpi_settings = df.dropna(subset=["Low Risk", "Moderate Risk", "Hig]()_
