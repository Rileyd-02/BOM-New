import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="SAP vs PLM Consumption Comparator", layout="wide")

st.title("SAP vs PLM Size-wise Consumption Comparison Tool")

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

    # Clean column names
    sap_df.columns = sap_df.columns.str.strip()
    plm_df.columns = plm_df.columns.str.strip()

    # -------------------------
    # Rename Columns Properly
    # -------------------------

    sap_df = sap_df.rename(columns={
        "Vendor Reference": "Vendor Ref",
        "Consumption": "SAP_Consumption",
        "Component": "Material",
        "Garment Size": "Garment Size"
    })

    plm_df = plm_df.rename(columns={
        "Consumption": "PLM_Consumption"
    })

    # -------------------------
    # Validate Required Columns
    # -------------------------

    required_sap_cols = ["Material", "Vendor Ref", "SAP_Consumption"]
    required_plm_cols = ["Material", "Vendor Ref", "PLM_Consumption", "Garment Size"]

    for col in required_sap_cols:
        if col not in sap_df.columns:
            st.error(f"Missing column in SAP file: {col}")
            st.stop()

    for col in required_plm_cols:
        if col not in plm_df.columns:
            st.error(f"Missing column in PLM file: {col}")
            st.stop()

    # -------------------------
    # Ensure Size Column Exists in SAP
    # -------------------------

    if "Garment Size" not in sap_df.columns:
        st.warning("Garment Size not found in SAP file. Size-wise comparison may not work properly.")

    # -------------------------
    # Convert to Decimal
    # -------------------------

    sap_df["SAP_Consumption"] = pd.to_numeric(
        sap_df["SAP_Consumption"], errors="coerce"
    ).fillna(0.0)

    plm_df["PLM_Consumption"] = pd.to_numeric(
        plm_df["PLM_Consumption"], errors="coerce"
    ).fillna(0.0)

    # -------------------------
    # Merge Size-wise
    # -------------------------

    merge_keys = ["Material", "Vendor Ref"]

    if "Garment Size" in sap_df.columns:
        merge_keys.append("Garment Size")

    merged_df = pd.merge(
        sap_df,
        plm_df,
        on=merge_keys,
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

    # Round for display only
    merged_df["SAP_Consumption"] = merged_df["SAP_Consumption"].round(4)
    merged_df["PLM_Consumption"] = merged_df["PLM_Consumption"].round(4)
    merged_df["Consumption_Difference"] = merged_df["Consumption_Difference"].round(4)
    merged_df["Consumption_Diff_%"] = merged_df["Consumption_Diff_%"].round(2)

    # -------------------------
    # Status Filter
    # -------------------------

    status_filter = st.selectbox(
        "Filter by Status",
        ["All"] + list(merged_df["Consumption_Status"].unique())
    )

    if status_filter != "All":
        final_df = merged_df[merged_df["Consumption_Status"] == status_filter]
    else:
        final_df = merged_df

    # -------------------------
    # Output Columns
    # -------------------------

    display_cols = [
        "Material",
        "Vendor Ref",
        "Garment Size",
        "SAP_Consumption",
        "PLM_Consumption",
        "Consumption_Difference",
        "Consumption_Diff_%",
        "Consumption_Status"
    ]

    display_cols = [col for col in display_cols if col in final_df.columns]

    final_df = final_df[display_cols]

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
        file_name="SAP_vs_PLM_Sizewise_Comparison.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
