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

# Drop duplicate columns, keep the first occurrence
df = df.loc[:, ~df.columns.duplicated()].copy()

# Rename the first 5 useful columns
df = df.rename(columns={
    df.columns[0]: "Metric",
    df.columns[1]: "Low Risk",
    df.columns[2]: "Moderate Risk",
    df.columns[3]: "High Risk",
    df.columns[4]: "Weight"
})

# --- Create Section column ---
# Section = Metric name when thresholds are blank, forward-fill otherwise
df['Section'] = df.apply(
    lambda row: row['Metric'] if pd.isna(row["Low Risk"]) else None, axis=1
)
df['Section'] = df['Section'].ffill()

# Keep only rows with thresholds + weights (actual KPIs)
kpi_settings = df.dropna(subset=["Low Risk", "Moderate Risk", "High Risk", "Weight"])

# --- Debug Info ---
st.write("DEBUG - Available KPI columns:", kpi_settings.columns.tolist())
s
