import streamlit as st
import pandas as pd
from io import BytesIO

# ------------------------
# Page Setup
# ------------------------
st.set_page_config(page_title="SAP vs PLM Consumption Validation", layout="wide")
st.title("📊 SAP vs PLM Consumption Validation Tool")

st.write("""
This tool validates **SAP vs PLM consumption in decimals**.

✔ SAP consumption is normalized using **Comp.Qty / Base Quantity**  
✔ PLM consumption is used **as-is (decimals)**  
✔ Match logic: **Material + Vendor Reference**  
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
        # Rename Required Columns
        # ------------------------
        sap_df.rename(columns={
            "Material": "Material",
            "Vendor Reference": "Vendor_Ref",
            "Comp.Qty.": "SAP_Comp_Qty",
            "Base quantity": "Base_Qty"
        }, inplace=True)

        plm_df.rename(columns={
            "Material": "Material",
            "Vendor Ref": "Vendor_Ref",
            "Consumption": "PLM_Consumption"
        }, inplace=True)

        # ------------------------
        # Validate Required Columns
        # ------------------------
        required_sap = ["Material", "Vendor_Ref", "SAP_Comp_Qty", "Base_Qty"]
        required_plm = ["Material", "Vendor_Ref", "PLM_Consumption"]

        for col in required_sap:
            if col not in sap_df.columns:
                st.error(f"❌ Missing column in SAP file: {col}")
                st.stop()

        for col in required_plm:
            if col not in plm_df.columns:
                st.error(f"❌ Missing column in PLM file: {col}")
                st.stop()

        # ------------------------
        # Clean & Normalize Join Keys (CRITICAL FIX)
        # ------------------------
        sap_df["Material"] = sap_df["Material"].astype(str).str.strip()
        plm_df["Material"] = plm_df["Material"].astype(str).str.strip()

        sap_df["Vendor_Ref"] = sap_df["Vendor_Ref"].astype(str).str.strip()
        plm_df["Vendor_Ref"] = plm_df["Vendor_Ref"].astype(str).str.strip()

        # ------------------------
        # Convert Numeric Columns
        # ------------------------
        sap_df["SAP_Comp_Qty"] = pd.to_numeric(sap_df["SAP_Comp_Qty"], errors="coerce")
        sap_df["Base_Qty"] = pd.to_numeric(sap_df["Base_Qty"], errors="coerce")
        plm_df["PLM_Consumption"] = pd.to_numeric(plm_df["PLM_Consumption"], errors="coerce")

        # ------------------------
        # SAP Consumption Normalization (DECIMALS)
        # ------------------------
        sap_df["SAP_Consumption"] = sap_df.apply(
            lambda x: round(x["SAP_Comp_Qty"] / x["Base_Qty"], 5)
            if pd.notna(x["SAP_Comp_Qty"]) and pd.notna(x["Base_Qty"]) and x["Base_Qty"] != 0
            else None,
            axis=1
        )

        # ------------------------
        # Merge SAP & PLM
        # ------------------------
        merged_df = pd.merge(
            sap_df,
            plm_df,
            on=["Material", "Vendor_Ref"],
            how="left"
        )

        # ------------------------
        # Consumption Comparison (DECIMAL SAFE)
        # ------------------------
        TOLERANCE = 0.001  # ~0.1%

        def consumption_status(row):
            sap = row["SAP_Consumption"]
            plm = row["PLM_Consumption"]

            if pd.isna(plm):
                return "Missing in PLM"
            if pd.isna(sap):
                return "Missing in SAP"

            diff = abs(sap - plm)

            if diff <= TOLERANCE:
                return "MATCH"
            elif sap > plm:
                return "SAP Higher"
            else:
                return "PLM Higher"

        merged_df["Consumption_Status"] = merged_df.apply(consumption_status, axis=1)

        merged_df["Consumption_Difference"] = merged_df.apply(
            lambda x: round(x["SAP_Consumption"] - x["PLM_Consumption"], 5)
            if pd.notna(x["SAP_Consumption"]) and pd.notna(x["PLM_Consumption"])
            else None,
            axis=1
        )

        # ------------------------
        # Summary
        # ------------------------
        st.subheader("📈 Summary")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Records", len(merged_df))
        c2.metric("Matches", (merged_df["Consumption_Status"] == "MATCH").sum())
        c3.metric("Mismatches", (merged_df["Consumption_Status"] != "MATCH").sum())

        # ------------------------
        # Preview
        # ------------------------
        st.subheader("🔍 Preview Results")
        st.dataframe(
            merged_df[
                [
                    "Material",
                    "Vendor_Ref",
                    "SAP_Comp_Qty",
                    "Base_Qty",
                    "SAP_Consumption",
                    "PLM_Consumption",
                    "Consumption_Difference",
                    "Consumption_Status"
                ]
            ].head(200)
        )

        # ------------------------
        # Excel Export
        # ------------------------
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            merged_df.to_excel(writer, sheet_name="Consumption_Comparison", index=False)

            workbook = writer.book
            worksheet = writer.sheets["Consumption_Comparison"]

            status_col = merged_df.columns.get_loc("Consumption_Status")

            green = workbook.add_format({'bg_color': '#C6EFCE'})
            red = workbook.add_format({'bg_color': '#FFC7CE'})
            orange = workbook.add_format({'bg_color': '#FFD580'})

            worksheet.conditional_format(
                1, status_col, len(merged_df), status_col,
                {'type': 'text', 'criteria': 'containing', 'value': 'MATCH', 'format': green}
            )
            worksheet.conditional_format(
                1, status_col, len(merged_df), status_col,
                {'type': 'text', 'criteria': 'containing', 'value': 'Higher', 'format': red}
            )
            worksheet.conditional_format(
                1, status_col, len(merged_df), status_col,
                {'type': 'text', 'criteria': 'containing', 'value': 'Missing', 'format': orange}
            )

        output.seek(0)

        st.download_button(
            label="📥 Download Comparison Output",
            data=output,
            file_name="SAP_PLM_Consumption_Comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error while processing: {e}")

else:
    st.info("⬆️ Please upload both SAP and PLM Excel files.")
