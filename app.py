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

# Ensure first column is always "Metric"
if df.columns[0] != "Metric":
    df = df.rename(columns={df.columns[0]: "Metric"})

# Define important columns
keep_cols = ["Metric", "Low Risk", "Moderate Risk", "High Risk", "Weight", "Section"]

# Keep only columns that actually exist
df = df[[c for c in keep_cols if c in df.columns]]

# Convert thresholds and weights to numeric
numeric_cols = [c for c in ["Low Risk", "Moderate Risk", "High Risk", "Weight"] if c in df.columns]
df[numeric_cols] = df[numeric_cols].apply(pd.to_numeric, errors="coerce")

# Ensure Section column exists and is forward-filled
if "Section" not in df.columns:
    df["Section"] = df["Metric"].where(df.get("Low Risk").isna()).ffill()

# Keep only rows with valid thresholds + weights
kpi_settings = df.dropna(subset=[c for c in ["Low Risk", "Moderate Risk", "High Risk", "Weight"] if c in df.columns])

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
        section_weight = 0

        for _, row in section_df.iterrows():
            metric = row['Metric']
            low = row['Low Risk']
            med = row['Moderate Risk']
            high = row['High Risk']
            weight = row['Weight']
            value = user_inputs[metric]

            # Risk scoring logic
            if pd.notna(low) and value >= low:
                score = 3
                level = "Low Risk"
            elif pd.notna(med) and value >= med:
                score = 2
                level = "Moderate Risk"
            else:
                score = 1
                level = "High Risk"

            weighted_score = score * weight
            total_score += weighted_score
            total_weight += weight

            section_score += weighted_score
            section_weight += weight

            detailed_results.append({
                "Section": section,
                "Metric": metric,
                "Value": value,
                "Risk Level": level,
                "Score": score,
                "Weight": weight,
                "Weighted Score": weighted_score
            })

        # Section average
        if section_weight > 0:
            section_results[section] = section_score / section_weight
        else:
            section_results[section] = 0

    # Final score
    final_score = total_score / total_weight if total_weight > 0 else 0

    # --- Display Results ---
    st.subheader("Section Scores")
    for section, score in section_results.items():
        st.metric(section, f"{score:.2f}")

    st.subheader("Final Customer Health Score")
    st.metric("Overall Score", f"{final_score:.2f}")

    if final_score >= 2.5:
        st.success("Status: GREEN (Healthy)")
    elif final_score >= 1.75:
        st.warning("Status: YELLOW (At Risk)")
    else:
        st.error("Status: RED (High Risk)")

    st.subheader("Detailed Results")
    st.write(pd.DataFrame(detailed_results))
