import streamlit as st
import pandas as pd

# --- Load Weights and Thresholds from Excel ---
excel_file = "LNE CustomerHealthScoringModel.xlsx"

# Load Input sheet without assuming headers
df = pd.read_excel(excel_file, sheet_name="Input", header=None)

# Debug: show first few rows
st.write("Preview of Input sheet:", df.head(20))

# For now, just use the whole dataframe until we align names
kpi_settings = df

# Expecting columns: Section, Metric, Low Risk, Moderate Risk, High Risk, Weight
kpi_settings = df[['Section', 'Metric', 'Low Risk', 'Moderate Risk', 'High Risk', 'Weight']]

# --- Streamlit App ---
st.title("Customer Lifecycle Management Health Score")
st.markdown("Enter your actual KPI metrics below to calculate section scores and the final weighted health score.")

user_inputs = {}

# Group inputs by section
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

    # Calculate section scores
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
            if value >= low:
                score = 3
                level = "Low Risk"
            elif value >= med:
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


