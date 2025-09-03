import streamlit as st
import pandas as pd

# --- Load Excel Config ---
excel_file = "LNE CustomerHealthScoringModel.xlsx"

# Load Input sheet raw (no assumption about headers yet)
raw_df = pd.read_excel(excel_file, sheet_name="Input", header=None)

# Try to detect header row (search for "Metric"), fallback to row 0
header_rows = raw_df[raw_df.apply(lambda r: r.astype(str).str.contains("Metric").any(), axis=1)].index
if len(header_rows) > 0:
    header_row = header_rows[0]
else:
    header_row = 0

# Re-load with proper headers
df = pd.read_excel(excel_file, sheet_name="Input", header=header_row)

# --- Clean & Align Columns ---
# Rename based on your CSV structure
df = df.rename(columns={
    df.columns[1]: "Metric",
    df.columns[2]: "Low Risk",
    df.columns[3]: "Moderate Risk",
    df.columns[4]: "High Risk",
    df.columns[5]: "CustomerInput",
    df.columns[6]: "Weight"
})

# Convert numeric fields
for col in ["Low Risk", "Moderate Risk", "High Risk", "CustomerInput", "Weight"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

# Create Section column: rows with Weight but no CustomerInput are headers
df["Section"] = df["Metric"].where(df["Weight"].notna() & df["CustomerInput"].isna()).ffill()

# KPI rows = rows with a Metric and Weight
kpi_settings = df[df["Metric"].notna() & df["Weight"].notna()]

# --- Debug Preview ---
st.write("DEBUG - Columns loaded:", df.columns.tolist())
st.write("DEBUG - KPI Settings Preview", kpi_settings.head(20))

# --- Streamlit App ---
st.title("Customer Lifecycle Management Health Score")
st.markdown("Enter your actual KPI metrics below to calculate section scores and the final weighted health score.")

user_inputs = {}

# KPI Input Form
with st.form("kpi_form"):
    for section in kpi_settings["Section"].unique():
        st.subheader(section)
        section_df = kpi_settings[kpi_settings["Section"] == section]
        for _, row in section_df.iterrows():
            metric = row["Metric"]
            default_val = float(row["CustomerInput"]) if pd.notna(row["CustomerInput"]) else 0.0
            value = st.number_input(f"{metric}", min_value=0.0, step=0.1, value=default_val)
            user_inputs[metric] = value
        st.markdown("---")
    submitted = st.form_submit_button("Calculate Score")

# --- Process Results ---
if submitted:
    section_results = {}
    total_score = 0
    total_weight = 0
    detailed_results = []

    for section in kpi_settings["Section"].unique():
        section_df = kpi_settings[kpi_settings["Section"] == section]
        section_score = 0
        section_weight = 0

        for _, row in section_df.iterrows():
            metric = row["Metric"]
            if metric not in user_inputs:
                continue

            low = row.get("Low Risk")
            med = row.get("Moderate Risk")
            high = row.get("High Risk")
            weight = row.get("Weight")
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

            weighted_score = score * weight if pd.notna(weight) else score
            total_score += weighted_score
            total_weight += weight if pd.notna(weight) else 1

            section_score += weighted_score
            section_weight += weight if pd.notna(weight) else 1

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
