import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="SAP vs PLM Consumption Validation", layout="wide")
st.title("📊 SAP vs PLM Consumption Validation Tool")

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
        # PREVIEW RAW COLUMNS
        # ------------------------
        st.subheader("🧾 SAP Columns")
        st.write(list(sap_df.columns))

        st.subheader("🧾 PLM Columns")
        st.write(list(plm_df.columns))

        # ------------------------
        # Rename Columns
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
        # PREVIEW AFTER RENAME
        # ------------------------
        st.subheader("🔎 SAP Preview (After Rename)")
        st.dataframe(sap_df[["Material", "Vendor_Ref", "SAP_Comp_Qty", "Base_Qty"]].head(10))

        st.subheader("🔎 PLM Preview (After Rename)")
        st.dataframe(plm_df[["Material", "Vendor_Ref", "PLM_Consumption"]].head(10))

        # ------------------------
        # NORMALIZE JOIN KEYS (CRITICAL)
        # ------------------------

        # Convert material to string & REMOVE leading zeros
        sap_df["Material"] = (
            sap_df["Material"]
            .astype(str)
            .str.strip()
            .str.lstrip("0")
        )

        plm_df["Material"] = (
            plm_df["Material"]
            .astype(str)
            .str.strip()
            .str.lstrip("0")
        )

        sap_df["Vendor_Ref"] = sap_df["Vendor_Ref"].astype(str).str.strip()
        plm_df["Vendor_Ref"] = plm_df["Vendor_Ref"].astype(str).str.strip()

        # ------------------------
        # DEBUG JOIN KEYS
        # ------------------------
        st.subheader("🧩 SAP Join Keys")
        st.dataframe(sap_df[["Material", "Vendor_Ref"]].drop_duplicates().head(10))

        st.subheader("🧩 PLM Join Keys")
        st.dataframe(plm_df[["Material", "Vendor_Ref"]].drop_duplicates().head(10))

        # ------------------------
        # Numeric Conversion
        # ------------------------
        sap_df["SAP_Comp_Qty"] = pd.to_numeric(sap_df["SAP_Comp_Qty"], errors="coerce")
        sap_df["Base_Qty"] = pd.to_numeric(sap_df["Base_Qty"], errors="coerce")
        plm_df["PLM_Consumption"] = pd.to_numeric(plm_df["PLM_Consumption"], errors="coerce")

        # ------------------------
        # SAP Consumption (DECIMAL)
        # ------------------------
        sap_df["SAP_Consumption"] = sap_df.apply(
            lambda x: round(x["SAP_Comp_Qty"] / x["Base_Qty"], 5)
            if pd.notna(x["SAP_Comp_Qty"]) and pd.notna(x["Base_Qty"]) and x["Base_Qty"] != 0
            else None,
            axis=1
        )

        # ------------------------
        # MERGE
        # ------------------------
        merged_df = pd.merge(
            sap_df,
            plm_df,
            on=["Material", "Vendor_Ref"],
            how="left"
        )

        # ------------------------
        # STATUS
        # ------------------------
        TOLERANCE = 0.001

        def status(row):
            if pd.isna(row["PLM_Consumption"]):
                return "Missing in PLM"
            if abs(row["SAP_Consumption"] - row["PLM_Consumption"]) <= TOLERANCE:
                return "MATCH"
            return "Mismatch"

        merged_df["Status"] = merged_df.apply(status, axis=1)

        merged_df["Difference"] = merged_df["SAP_Consumption"] - merged_df["PLM_Consumption"]

        # ------------------------
        # FINAL PREVIEW
        # ------------------------
        st.subheader("✅ Final Comparison Preview")
        st.dataframe(
            merged_df[
                [
                    "Material",
                    "Vendor_Ref",
                    "SAP_Comp_Qty",
                    "Base_Qty",
                    "SAP_Consumption",
                    "PLM_Consumption",
                    "Difference",
                    "Status"
                ]
            ].head(200)
        )

        # ------------------------
        # EXPORT
        # ------------------------
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            merged_df.to_excel(writer, index=False, sheet_name="Comparison")

        output.seek(0)

        st.download_button(
            "📥 Download Output",
            data=output,
            file_name="SAP_PLM_Consumption_Comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")

else:
    st.info("⬆️ Upload both SAP and PLM files to begin.")
