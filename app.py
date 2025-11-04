import streamlit as st
import pandas as pd

st.set_page_config(page_title="Production Report Summary", layout="wide")

st.title("🛢️ NORPETCO Production Report Summary")

# File uploader
uploaded_file = st.file_uploader("📂 Upload the Excel Report (.xlsm or .xlsx)", type=["xlsm", "xlsx"])

if uploaded_file:
    try:
        # Read the Excel "Report" sheet
        df = pd.read_excel(uploaded_file, sheet_name="Report", header=0)

        # Clean column names
        df.columns = df.columns.astype(str).str.strip()

        # Show columns for verification
        st.write("🧩 Detected columns in Excel:", list(df.columns))

        # Try to automatically find matching column names
        well_col = next((c for c in df.columns if "running" in c.lower()), None)
        net_bo_col = next((c for c in df.columns if "net" in c.lower() and "bo" in c.lower()), None)
        net_diff_col = next((c for c in df.columns if "net" in c.lower() and "diff" in c.lower()), None)
        wc_col = next((c for c in df.columns if "w/c" in c.lower() or "wc" in c.lower()), None)
        zone_col = next((c for c in df.columns if "zone" in c.lower()), None)

        # Check columns
        st.write("✅ Matched columns:", {
            "Well Name": well_col,
            "Production Zone": zone_col,
            "Net BO": net_bo_col,
            "Net Diff BO": net_diff_col,
            "W/C": wc_col
        })

        if not all([well_col, net_bo_col, net_diff_col, wc_col]):
            st.error("❌ Could not automatically detect required columns. Please check your Excel header names.")
            st.stop()

        # Select and rename columns
        wells_df = df[[well_col, zone_col, net_bo_col, net_diff_col, wc_col]].copy()
        wells_df.columns = ["Well Name", "Production Zone", "TOTAL PRODUCTION", "NET DIFF", "W/C"]

        # Convert to numeric
        wells_df["TOTAL PRODUCTION"] = pd.to_numeric(wells_df["TOTAL PRODUCTION"], errors="coerce")
        wells_df["NET DIFF"] = pd.to_numeric(wells_df["NET DIFF"], errors="coerce")

        # Remove blank rows
        wells_df = wells_df.dropna(subset=["Well Name", "TOTAL PRODUCTION"], how="any")

        # Compute totals
        total_production = wells_df["TOTAL PRODUCTION"].sum()
        total_net_diff = wells_df["NET DIFF"].sum()

        # Add total row
        total_row = pd.DataFrame({
            "Well Name": ["TOTAL"],
            "Production Zone": [""],
            "TOTAL PRODUCTION": [total_production],
            "NET DIFF": [total_net_diff],
            "W/C": [""]
        })
        wells_df = pd.concat([wells_df, total_row], ignore_index=True)

        # Display styled table
        st.dataframe(
            wells_df.style.format({
                "TOTAL PRODUCTION": "{:,.0f}",
                "NET DIFF": "{:+.0f}",
                "W/C": "{}"
            }),
            use_container_width=True
        )

        st.success("✅ Table generated successfully!")

    except Exception as e:
        st.error(f"⚠️ Error reading Excel file: {e}")

else:
    st.info("📥 Please upload your Excel production report to continue.")
