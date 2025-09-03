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
df = df.loc[:, ~df.columns.duplicated()].copy()

# Ensure first column is always "Metric"
if df.columns[0] != "Metric":
    df = df.rename(columns={df.columns[0]: "Metric"})

# Convert thresholds and weights to numeric (force, allow NaN)
for col in ["Low Risk", "Moderate Risk", "High Risk", "Weight"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

# Ensure Section column exists and is forward-f
