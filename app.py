import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="SAP vs PLM Consumption Comparator", layout="wide")

st.title("SAP vs PLM Consumption Comparison Tool")

st.markdown("Upload SAP and PLM files to compare consumption with decimal precision.")

# -------------------------
# File Upload
# -------------------------

sap_file = st.file_uploader("Upload SAP File", type=["xlsx"])
plm_file = st.file_uploader("Upload PLM File", type=["xlsx"])

tolerance_percent = st.number_input(
    "Tolerance (%)",
    min_value=0.0,
    max_value=100.0,
    value=5.0,
    step=0.1
)

if sap_file and plm_file:

    # -------------------------
    # Load Files
    # -------------------------

    sap_df = pd.read_excel(sap_file)
    plm_df = pd.read_excel(plm_file)

    # -------------------------
    # Rename Columns (Adjust if needed)
    # -------------------------

    sap_df = sap_df.rename(columns={
        "Component": "Material",
        "Component Qty": "SAP_Component_Qty",
        "Base Qty": "Base_Qty",
        "Vendor Ref": "Vendor Ref"
    })

    plm_df = plm_df.rename(columns={
        "Material": "Material",
        "Consumption": "PLM_Consumption",
        "Vendor Ref": "Vendor Ref",
        "Garment Size": "Garment Size"
    })

    # -------------------------
    # Convert to Decimal Safe Numeric
    # -------------------------

    sap_df["SAP_Component_Qty"] = pd.to_numeric(sap_df["SAP_Component_Qty"], errors="coerce")
    sap_df["Base_Qty"] = pd.to_numeric(sap_df["Base_Qty"], errors="coerce")
    plm_df["PLM_Consumption"] = pd.to_numeric(plm_df["PLM_Consumption"], errors="coerce")

    # -------------------------
    # Calculate SAP Consumption
    # -------------------------

    sap_df["SAP_Consumption"] = sap_df.apply(
        lambda x: x["SAP_Component_Qty"] / x["Base_Qty"]
        if pd.notna(x["Base_Qty"]) and x["Base_Qty"] != 0
        else 0.0,
        axis=1
    )

    # -------------------------
    # Merge SAP + PLM
    # -------------------------

    merged_df = pd.merge(
        sap_df,
        plm_df,
        on=["Material", "Vendor Ref"],
        how="outer"
    )

    merged_df["SAP_Consumption"] = merged_df["SAP_Consumption"].fillna(0.0)
    merged_df["PLM_Consumption"] = merged_df["PLM_Consumption"].fillna(0.0)

    # -------------------------
    # Calculate Difference
    # -------------------------

    merged_df["Consumption_Difference"] = (
        merged_df["SAP_Consumption"] - merged_df["PLM_Consumption"]
    )

    # Percentage difference (based on PLM)
    merged_df["Consumption_Diff_%"] = merged_df.apply(
        lambda x: ((x["SAP_Consumption"] - x["PLM_Consumption"]) / x["PLM_Consumption"]) * 100
        if x["PLM_Consumption"] != 0
        else None,
        axis=1
    )

    # -------------------------
    # Tolerance Logic
    # -------------------------

    tolerance = tolerance_percent / 100

    def compare_consumption(row):
        sap = row["SAP_Consumption"]
        plm = row["PLM_Consumption"]

        if plm == 0 and sap == 0:
            return "OK"
        if plm == 0 and sap != 0:
            return "PLM Missing"
        if sap == 0 and plm != 0:
            return "SAP Missing"

        diff_pct = abs(sap - plm) / plm

        if diff_pct <= tolerance:
            return "Within Tolerance"
        elif sap > plm:
            return "SAP Higher"
        else:
            return "PLM Higher"

    merged_df["Consumption_Status"] = merged_df.apply(compare_consumption, axis=1)

    # -------------------------
    # Round for Display Only
    # -------------------------

    merged_df["SAP_Consumption"] = merged_df["SAP_Consumption"].round(4)
    merged_df["PLM_Consumption"] = merged_df["PLM_Consumption"].round(4)
    merged_df["Consumption_Difference"] = merged_df["Consumption_Difference"].round(4)
    merged_df["Consumption_Diff_%"] = merged_df["Consumption_Diff_%"].round(2)

    # -------------------------
    # Filter Option
    # -------------------------

    status_filter = st.selectbox(
        "Filter by Status",
        ["All"] + list(merged_df["Consumption_Status"].unique())
    )

    if status_filter != "All":
        filtered_df = merged_df[merged_df["Consumption_Status"] == status_filter]
    else:
        filtered_df = merged_df

    # -------------------------
    # Select Output Columns
    # -------------------------

    final_df = filtered_df[[
        "Material",
        "Vendor Ref",
        "Garment Size",
        "SAP_Consumption",
        "PLM_Consumption",
        "Consumption_Difference",
        "Consumption_Diff_%",
        "Consumption_Status"
    ]]

    st.dataframe(final_df, use_container_width=True)

    # -------------------------
    # Download Excel
    # -------------------------

    def to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Comparison")
        return output.getvalue()

    st.download_button(
        label="Download Comparison Report",
        data=to_excel(final_df),
        file_name="SAP_vs_PLM_Comparison.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
