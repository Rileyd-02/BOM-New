import streamlit as st
import pandas as pd
from io import BytesIO

# ------------------------
# Page Setup
# ------------------------
st.set_page_config(page_title="SAP vs PLM Consumption Validation", layout="wide")
st.title("📊 SAP vs PLM Consumption Comparison")

st.write("""
This tool:
- Normalizes SAP consumption using **Comp.Qty / Base quantity**
- Compares with **PLM decimal consumption**
- Matches by **Material + Vendor Reference**
- Outputs a clean Excel comparison
""")

# ------------------------
# File Upload
# ------------------------
sap_file = st.file_uploader("📤 Upload SAP Excel File", type=["xlsx"])
plm_file = st.file_uploader("📤 Upload PLM Excel File", type=["xlsx"])

if sap_file and plm_file:
    try:
        # ------------------------
        # Read Files
        # ------------------------
        sap_df = pd.read_excel(sap_file)
        plm_df = pd.read_excel(plm_file)

        sap_df.columns = sap_df.columns.str.strip()
        plm_df.columns = plm_df.columns.str.strip()

        # ------------------------
        # Column Mapping
        # ------------------------
        sap_df = sap_df.rename(columns={
            "Material": "Material",
            "Vendor Reference": "Vendor_Ref",
            "Comp.Qty": "SAP_Comp_Qty",
            "Base quantity": "Base_Qty"
        })

        plm_df = plm_df.rename(columns={
            "Material": "Material",
            "Vendor Ref": "Vendor_Ref",
            "Consumption": "PLM_Consumption"
        })

        # ------------------------
        # Convert to Numeric (Decimals)
        # ------------------------
        sap_df["SAP_Comp_Qty"] = pd.to_numeric(sap_df["SAP_Comp_Qty"], errors="coerce")
        sap_df["Base_Qty"] = pd.to_numeric(sap_df["Base_Qty"], errors="coerce")
        plm_df["PLM_Consumption"] = pd.to_numeric(plm_df["PLM_Consumption"], errors="coerce")

        # ------------------------
        # Normalize SAP Consumption
        # ------------------------
        sap_df["SAP_Consumption"] = (
            sap_df["SAP_Comp_Qty"] / sap_df["Base_Qty"]
        ).round(5)

        # ------------------------
        # Merge SAP & PLM
        # ------------------------
        merged_df = pd.merge(
            sap_df,
            plm_df,
            on=["Material", "Vendor_Ref"],
            how="left"
        )

        merged_df["PLM_Consumption"] = merged_df["PLM_Consumption"].round(5)

        # ------------------------
        # Consumption Comparison
        # ------------------------
        def compare_consumption(row):
            sap = row["SAP_Consumption"]
            plm = row["PLM_Consumption"]

            if pd.isna(plm):
                return "Missing in PLM"
            if pd.isna(sap):
                return "Missing in SAP"

            if round(sap, 5) == round(plm, 5):
                return "MATCH"
            elif sap > plm:
                return "SAP Higher"
            else:
                return "PLM Higher"

        merged_df["Consumption_Status"] = merged_df.apply(compare_consumption, axis=1)

        merged_df["Consumption_Diff"] = (
            merged_df["SAP_Consumption"] - merged_df["PLM_Consumption"]
        ).round(5)

        # ------------------------
        # Display Preview
        # ------------------------
        st.subheader("🔍 Comparison Preview")

        st.dataframe(
            merged_df[
                [
                    "Material",
                    "Vendor_Ref",
                    "SAP_Comp_Qty",
                    "Base_Qty",
                    "SAP_Consumption",
                    "PLM_Consumption",
                    "Consumption_Diff",
                    "Consumption_Status"
                ]
            ]
        )

        # ------------------------
        # Excel Export
        # ------------------------
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            merged_df.to_excel(writer, index=False, sheet_name="Consumption_Comparison")

            workbook = writer.book
            worksheet = writer.sheets["Consumption_Comparison"]

            status_col = merged_df.columns.get_loc("Consumption_Status")

            green = workbook.add_format({"bg_color": "#C6EFCE"})
            red = workbook.add_format({"bg_color": "#FFC7CE"})
            orange = workbook.add_format({"bg_color": "#FFD580"})

            worksheet.conditional_format(
                1, status_col, len(merged_df), status_col,
                {"type": "text", "criteria": "containing", "value": "MATCH", "format": green}
            )
            worksheet.conditional_format(
                1, status_col, len(merged_df), status_col,
                {"type": "text", "criteria": "containing", "value": "SAP Higher", "format": red}
            )
            worksheet.conditional_format(
                1, status_col, len(merged_df), status_col,
                {"type": "text", "criteria": "containing", "value": "PLM Higher", "format": orange}
            )

        output.seek(0)

        st.download_button(
            "📥 Download Comparison Output",
            data=output,
            file_name="SAP_vs_PLM_Consumption.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")

else:
    st.info("⬆️ Upload both SAP and PLM files to begin.")
