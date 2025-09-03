import streamlit as st
import pandas as pd

# --- Load Excel Config ---
excel_file = "LNE CustomerHealthScoringModel.xlsx"

# Load Input sheet, but don't assume headers are row 0
raw_df = pd.read_excel(excel_file, sheet_name="Input", header=None)

# Find the row that contains the true headers (look for 'Low Risk')
header_row = raw_df[raw_df.apply(lambda r: r.astype(str).str.contains("Low Risk").any(), axis=1)].index[0]

# Re-load with proper headers
df = pd.read_excel(excel_file, sheet_name="Input", header=header_row)

# Rename relevant columns (adjust to match your Excel exactly)
df = df.rename(columns={
    df.columns[0]: "Metric",
    df.columns[2]: "Low Risk",
    df.columns[3]: "Moderate Risk",
    df.columns[4]: "High Risk",
    df.columns[7]: "Weight"
})

# Add Section column by forward-filling metrics that look like section headers
df['Section'] = df['Metric']
df['Section'] = df['Section'].where(df['Low Risk'].notna(), None)  # blank thresholds = section name
df['Section'] = df['Section'].ffill()

# Keep only rows that have thresholds + weight (i.e., actual KPIs)
kpi_settings = df.dropna(subset=["Low Risk", "Moderate Risk", "High Risk", "Weight"])

# --- Streamlit App ---
st.title("Customer Lifecycle Management Health Score")
st.markdown("Enter your actual KPI metrics below to calculate section scores and the final weighted health score.")

user_inputs = {}

# KPI Input Form
with st.form("kpi_form"):
    for section in kpi_settings['Section'].unique():
        st.subheader(section)
        section_df = kpi_settings[kpi_settings['Section'] == section]
        for _, row in section_df.iterrows():
            metric = row['Metric']
            value = st.number_input(f"{metric}", min_value=0.0, step=0.1)
            user_inputs[metric] = value
    submitted = st.form_submit_button("Calculate Score")

# --- Process Results ---
if submitted:
    section_results = {}
    total_score = 0
    total_weight = 0
    detailed_results = []

    for section in kpi_settings['Section'].unique():
        section_df = kpi_settings[kpi_settings['Section'] == section]
        section_score = 0
        section_weight = 0_
