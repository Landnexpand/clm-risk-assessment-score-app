import streamlit as st
import pandas as pd

# --- Load Excel ---
excel_file = "LNE CustomerHealthScoringModel v2.xlsx"
df = pd.read_excel(excel_file, sheet_name="Input")

# --- Clean column names just in case ---
df.columns = df.columns.str.strip()

# --- Build Section column dynamically ---
def is_section_row(row):
    no_thresholds = pd.isna(row["Low Risk"]) and pd.isna(row["Moderate Risk"]) and pd.isna(row["High Risk"])
    no_weight = pd.isna(row["Weight"])
    return pd.notna(row["KPIs"]) and no_thresholds and no_weight

df["Section"] = df.apply(lambda r: r["KPIs"] if is_section_row(r) else None, axis=1).ffill()

# --- UI Header ---
st.title("Customer Lifecycle Management Health Score")
st.markdown("Compare your KPIs against thresholds and enter your actual values to calculate scores.")

user_inputs = {}
results = []

# --- Input + Scoring Loop ---
for i, row in df.iterrows():
    kpi = row["KPIs"]
    low, med, high = row["Low Risk"], row["Moderate Risk"], row["High Risk"]
    weight = row["Weight"]
    max_score = row["Max CLM Score"]
    section = row["Section"]

    # Skip pure section rows (no thresholds/weights)
    if is_section_row(row):
        continue

    # Layout: KPI | Low | Med | High | Input
    cols = st.columns([3, 2, 2, 2, 2])
    cols[0].write(kpi)
    cols[1].write(low if pd.notna(low) else "")
    cols[2].write(med if pd.notna(med) else "")
    cols[3].write(high if pd.notna(high) else "")

    # Input field
    val = cols[4].number_input(
        f"Input_{i}", 
        min_value=0.0, 
        step=0.1, 
        value=0.0, 
        label_visibility="collapsed"
    )
    user_inputs[kpi] = val

    # ---- Scoring Algorithm ----
    if pd.notna(low) and val >= low:
        clm_score = 3
    elif pd.notna(med) and val >= med:
        clm_score = 2
    else:
        clm_score = 1

    customer_score = clm_score * weight if pd.notna(weight) else clm_score

    results.append({
        "Section": section,
        "KPI": kpi,
        "Input": val,
        "CLM Score": clm_score,
        "Weight": weight,
        "Customer CLM Score": customer_score,
        "Max CLM Score": max_score
    })

# --- Convert to DataFrame ---
res_df = pd.DataFrame(results)

# --- Section Aggregation ---
section_summary = res_df.groupby("Section").agg({
    "Customer CLM Score": "sum",
    "Max CLM Score": "sum"
})
section_summary["CLM % Score"] = section_summary["Customer CLM Score"] / section_summary["Max CLM Score"]

# --- Overall Score ---
total_customer = section_summary["Customer CLM Score"].sum()
total_max = section_summary["Max CLM Score"].sum()
overall_pct = total_customer / total_max if total_max > 0 else 0

# --- Display Results ---
st.subheader("Detailed KPI Results")
st.dataframe(res_df, use_container_width=True)

st.subheader("Section Summary")
st.dataframe(section_summary, use_container_width=True)

st.subheader("Overall Health Score")
st.metric("Final CLM % Score", f"{overall_pct:.2%}")
