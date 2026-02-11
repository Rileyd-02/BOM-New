import streamlit as st
import pandas as pd
from thefuzz import fuzz
from io import BytesIO


st.set_page_config(page_title="SAP vs PLM Comparison", layout="wide")
st.title("📊 SAP vs PLM Validation Tool")

st.write("""
Upload your **SAP** and **PLM** files.  
This tool compares:

- Material + Color match  
- Vendor Reference match  
- Normalized SAP vs PLM consumption  
- Tolerance-based consumption differences  
- Fuzzy similarity (Vendor & Color)  
""")

# ------------------------
# File Upload
# ------------------------
sap_file = st.file_uploader("📤 Upload SAP Excel File", type=["xlsx"])
plm_file = st.file_uploader("📤 Upload PLM Excel File", type=["xlsx"])

if sap_file and plm_file:
    try:
        sap_df = pd.read_excel(sap_file)
        plm_df = pd.read_excel(plm_file)

        sap_df.columns = sap_df.columns.str.strip()
        plm_df.columns = plm_df.columns.str.strip()

        # ------------------------
        # Rename Columns
        # ------------------------
        sap_df.rename(columns={
            "Material": "Material",
            "FG Color Description": "Color",
            "Vendor Reference": "Vendor Reference_SAP",
            "Comp.Qty.": "SAP_Component_Qty",
            "Base quantity": "Base_Qty"
        }, inplace=True)

        plm_df.rename(columns={
            "Material": "Material",
            "Color Name": "Color",
            "Vendor Ref": "Vendor Reference_PLM",
            "Consumption": "PLM_Consumption"
        }, inplace=True)

        # ------------------------
        # Merge SAP & PLM
        # ------------------------
        merged_df = pd.merge(
            sap_df,
            plm_df,
            on=["Material", "Color"],
            how="left",
            suffixes=("_SAP", "_PLM")
        )

        merged_df["Material_Match"] = merged_df["PLM_Consumption"].apply(
            lambda x: "Matched in PLM" if pd.notna(x) else "Missing in PLM"
        )

        # ------------------------
        # Vendor Reference Check
        # ------------------------
        def check_vendor_ref(row):
            sap_ref = str(row.get("Vendor Reference_SAP", "")).strip()
            plm_ref = str(row.get("Vendor Reference_PLM", "")).strip()

            if not plm_ref:
                return "No Vendor Ref in PLM"
            if sap_ref == plm_ref:
                return "Exact Match"
            if sap_ref in plm_ref or plm_ref in sap_ref:
                return "Partial Match"
            return "Mismatch"

        merged_df["VendorRef_Status"] = merged_df.apply(check_vendor_ref, axis=1)

        # ------------------------
        # SAP Consumption Normalization
        # ------------------------
        merged_df["SAP_Consumption"] = merged_df.apply(
            lambda x: round(x["SAP_Component_Qty"] / x["Base_Qty"], 5)
            if pd.notna(x["Base_Qty"]) and x["Base_Qty"] != 0 else 0,
            axis=1
        )

        merged_df["PLM_Consumption"] = merged_df["PLM_Consumption"].fillna(0).round(5)

        # ------------------------
        # Consumption Comparison with Tolerance
        # ------------------------
        TOLERANCE = 0.05  # 5%

        def compare_consumption(row):
            sap = row["SAP_Consumption"]
            plm = row["PLM_Consumption"]

            if plm == 0 and sap == 0:
                return "OK"
            if plm == 0 and sap != 0:
                return "PLM Missing"

            diff_pct = abs(sap - plm) / plm

            if diff_pct <= TOLERANCE:
                return "Within Tolerance"
            elif sap > plm:
                return "SAP Higher"
            else:
                return "PLM Higher"

        merged_df["Consumption_Status"] = merged_df.apply(compare_consumption, axis=1)

        merged_df["Consumption_Diff_%"] = merged_df.apply(
            lambda x: round(((x["SAP_Consumption"] - x["PLM_Consumption"]) / x["PLM_Consumption"]) * 100, 2)
            if x["PLM_Consumption"] not in [0, None] else None,
            axis=1
        )

        # ------------------------
        # Fuzzy Similarity
        # ------------------------
        def smart_similarity(a, b):
            a, b = str(a).strip(), str(b).strip()
            if not a or not b:
                return 0
            return max(
                fuzz.token_sort_ratio(a, b),
                fuzz.token_set_ratio(a, b),
                fuzz.partial_ratio(a, b)
            )

        merged_df["Vendor_Similarity"] = merged_df.apply(
            lambda x: smart_similarity(x.get("Vendor Reference_SAP", ""), x.get("Vendor Reference_PLM", "")),
            axis=1
        )

        merged_df["Color_Similarity"] = merged_df.apply(
            lambda x: smart_similarity(x.get("Color_SAP", ""), x.get("Color_PLM", "")),
            axis=1
        )

        # ------------------------
        # FILTER OPTION
        # ------------------------
        st.subheader("🔎 Filter Results")
        show_mismatch_only = st.checkbox("Show only mismatches")

        if show_mismatch_only:
            filtered_df = merged_df[
                (merged_df["Consumption_Status"] != "Within Tolerance") |
                (merged_df["VendorRef_Status"] != "Exact Match") |
                (merged_df["Material_Match"] != "Matched in PLM")
            ]
        else:
            filtered_df = merged_df

        # ------------------------
        # Summary Metrics
        # ------------------------
        summary_df = merged_df.drop_duplicates(subset=["Material", "Color"])
        st.subheader("📈 Summary")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total FG Materials", len(summary_df))
        c2.metric("Within Tolerance", (summary_df["Consumption_Status"] == "Within Tolerance").sum())
        c3.metric("Mismatches", (summary_df["Consumption_Status"] != "Within Tolerance").sum())

        # ------------------------
        # Excel Export with Colors
        # ------------------------
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            merged_df.to_excel(writer, sheet_name="Comparison_Report", index=False)
            summary_df.to_excel(writer, sheet_name="Summary", index=False)

            workbook = writer.book
            worksheet = writer.sheets["Comparison_Report"]
            status_col_idx = merged_df.columns.get_loc("Consumption_Status")

            format_red = workbook.add_format({'bg_color': '#FFC7CE'})
            format_green = workbook.add_format({'bg_color': '#C6EFCE'})
            format_orange = workbook.add_format({'bg_color': '#FFD580'})

            worksheet.conditional_format(1, status_col_idx, len(merged_df), status_col_idx,
                {'type': 'text', 'criteria': 'containing', 'value': 'SAP Higher', 'format': format_red})
            worksheet.conditional_format(1, status_col_idx, len(merged_df), status_col_idx,
                {'type': 'text', 'criteria': 'containing', 'value': 'PLM Higher', 'format': format_orange})
            worksheet.conditional_format(1, status_col_idx, len(merged_df), status_col_idx,
                {'type': 'text', 'criteria': 'containing', 'value': 'Within Tolerance', 'format': format_green})

        output.seek(0)

        st.download_button(
            label="📥 Download Full Comparison Report",
            data=output,
            file_name="SAP_PLM_Comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # ------------------------
        # Preview Table
        # ------------------------
        st.subheader("🔍 Preview Results")
        preview_cols = [
            "Material", "Color",
            "Vendor Reference_SAP", "Vendor Reference_PLM", "VendorRef_Status",
            "SAP_Consumption", "PLM_Consumption",
            "Consumption_Diff_%", "Consumption_Status",
            "Vendor_Similarity", "Color_Similarity",
            "Material_Match"
        ]
        available_cols = [c for c in preview_cols if c in filtered_df.columns]
        st.dataframe(filtered_df[available_cols].head(200))

    except Exception as e:
        st.error(f"❌ Error while processing: {e}")

else:
    st.info("⬆️ Please upload both SAP and PLM files to start comparison.")
