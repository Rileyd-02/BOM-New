import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="SAP vs PLM Consumption Comparison", layout="wide")
st.title("SAP vs PLM Consumption Comparison (Decimal Safe)")

sap_file = st.file_uploader("Upload SAP File", type=["xlsx"])
plm_file = st.file_uploader("Upload PLM File", type=["xlsx"])

tolerance_percent = st.number_input(
    "Tolerance (%)",
    min_value=0.0,
    max_value=100.0,
    value=5.0,
    step=0.1
)

# ---- Fix duplicate columns (SAP has many duplicates)
def make_unique_columns(df):
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        idx = cols[cols == dup].index.tolist()
        for i, index in enumerate(idx):
            if i != 0:
                cols[index] = f"{dup}_{i}"
    df.columns = cols
    return df


if sap_file and plm_file:

    # ------------------------
    # Load files
    # ------------------------
    sap_df = pd.read_excel(sap_file)
    plm_df = pd.read_excel(plm_file)

    sap_df.columns = sap_df.columns.str.strip()
    plm_df.columns = plm_df.columns.str.strip()

    sap_df = make_unique_columns(sap_df)
    plm_df = make_unique_columns(plm_df)

    # ------------------------
    # Rename columns to common names
    # ------------------------
    sap_df = sap_df.rename(columns={
        "Component": "Material",
        "Vendor Reference": "Vendor Ref",
        "Consumption": "SAP_Consumption"
    })

    plm_df = plm_df.rename(columns={
        "Consumption": "PLM_Consumption"
    })

    # ------------------------
    # Validate columns
    # ------------------------
    for col in ["Material", "Vendor Ref", "SAP_Consumption"]:
        if col not in sap_df.columns:
            st.error(f"Missing column in SAP file: {col}")
            st.stop()

    for col in ["Material", "Vendor Ref", "PLM_Consumption"]:
        if col not in plm_df.columns:
            st.error(f"Missing column in PLM file: {col}")
            st.stop()

    # ------------------------
    # Clean merge keys
    # ------------------------
    for key in ["Material", "Vendor Ref"]:
        sap_df[key] = sap_df[key].astype(str).str.strip()
        plm_df[key] = plm_df[key].astype(str).str.strip()

    # ------------------------
    # Convert consumption to decimals
    # ------------------------
    sap_df["SAP_Consumption"] = pd.to_numeric(
        sap_df["SAP_Consumption"], errors="coerce"
    ).fillna(0.0)

    plm_df["PLM_Consumption"] = pd.to_numeric(
        plm_df["PLM_Consumption"], errors="coerce"
    ).fillna(0.0)

    # ------------------------
    # Merge SAP & PLM
    # ------------------------
    merged_df = pd.merge(
        sap_df,
        plm_df,
        on=["Material", "Vendor Ref"],
        how="outer"
    )

    merged_df["SAP_Consumption"] = merged_df["SAP_Consumption"].fillna(0.0)
    merged_df["PLM_Consumption"] = merged_df["PLM_Consumption"].fillna(0.0)

    # ------------------------
    # Calculate differences
    # ------------------------
    merged_df["Consumption_Difference"] = (
        merged_df["SAP_Consumption"] - merged_df["PLM_Consumption"]
    )

    merged_df["Consumption_Diff_%"] = merged_df.apply(
        lambda x: ((x["SAP_Consumption"] - x["PLM_Consumption"]) / x["PLM_Consumption"]) * 100
        if x["PLM_Consumption"] != 0
        else None,
        axis=1
    )

    # ------------------------
    # Tolerance logic
    # ------------------------
    tolerance = tolerance_percent / 100

    def status(row):
        sap = row["SAP_Consumption"]
        plm = row["PLM_Consumption"]

        if sap == 0 and plm == 0:
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

    merged_df["Consumption_Status"] = merged_df.apply(status, axis=1)

    # ------------------------
    # Round for display
    # ------------------------
    merged_df["SAP_Consumption"] = merged_df["SAP_Consumption"].round(4)
    merged_df["PLM_Consumption"] = merged_df["PLM_Consumption"].round(4)
    merged_df["Consumption_Difference"] = merged_df["Consumption_Difference"].round(4)
    merged_df["Consumption_Diff_%"] = merged_df["Consumption_Diff_%"].round(2)

    final_df = merged_df[
        [
            "Material",
            "Vendor Ref",
            "SAP_Consumption",
            "PLM_Consumption",
            "Consumption_Difference",
            "Consumption_Diff_%",
            "Consumption_Status"
        ]
    ]

    st.dataframe(final_df, use_container_width=True)

    # ------------------------
    # Download
    # ------------------------
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        final_df.to_excel(writer, index=False, sheet_name="Consumption_Comparison")

    st.download_button(
        "Download Excel Report",
        output.getvalue(),
        "SAP_vs_PLM_Consumption_Comparison.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
