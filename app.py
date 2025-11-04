import streamlit as st
import pandas as pd

st.set_page_config(page_title="Production Report Summary", layout="wide")

st.title("🛢️ NORPETCO Production Report Summary")

uploaded_file = st.file_uploader("📂 Upload the Excel Report (.xlsm or .xlsx)", type=["xlsm", "xlsx"])

if uploaded_file:
    try:
        # Read entire sheet (no header yet)
        df_raw = pd.read_excel(uploaded_file, sheet_name="Report", header=None)

        # Find start of the well data (look for "RUNNING WELLS")
        start_row = None
        for i, row in df_raw.iterrows():
            if row.astype(str).str.contains("RUNNING WELLS", case=False, na=False).any():
                start_row = i
                break

        if start_row is None:
            st.error("❌ Could not find 'RUNNING WELLS' in the sheet.")
            st.stop()

        # Extract table starting from RUNNING WELLS row
        df = pd.read_excel(uploaded_file, sheet_name="Report", header=start_row)

        # Clean up column names
        df.columns = df.columns.astype(str).str.strip()

        # Detect well name column (contains "WELL" or "RUNNING")
        well_col = next((c for c in df.columns if "well" in c.lower()), None)
        zone_col = next((c for c in df.columns if "zone" in c.lower()), None)
        net_bo_col = next((c for c in df.columns if "net" in c.lower() and "bo" in c.lower()), None)
        net_diff_col = next((c for c in df.columns if "net" in c.lower() and "diff" in c.lower()), None)
        wc_col = next((c for c in df.columns if "w/c" in c.lower() or "wc" in c.lower()), None)

        st.write("✅ Detected columns:", {
            "Well Name": well_col,
            "Zone": zone_col,
            "Net BO": net_bo_col,
            "Net Diff": net_diff_col,
            "W/C": wc_col
        })

        # Keep only the needed columns
        wells_df = df[[well_col, zone_col, net_bo_col, net_diff_col, wc_col]].copy()

        # Rename columns
        wells_df.columns = ["Well Name", "Production Zone", "TOTAL PRODUCTION", "NET DIFF", "W/C"]

        # Drop rows without well names
        wells_df = wells_df.dropna(subset=["Well Name"], how="any")

        # Stop before TOTAL row if present
        if (wells_df["Well Name"].astype(str).str.contains("TOTAL", case=False)).any():
            total_index = wells_df[wells_df["Well Name"].astype(str).str.contains("TOTAL", case=False)].index[0]
            wells_df = wells_df.loc[:total_index - 1]

        # Convert numeric columns
        wells_df["TOTAL PRODUCTION"] = pd.to_numeric(wells_df["TOTAL PRODUCTION"], errors="coerce")
        wells_df["NET DIFF"] = pd.to_numeric(wells_df["NET DIFF"], errors="coerce")

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

        # Display table
        st.dataframe(
            wells_df.style.format({
                "TOTAL PRODUCTION": "{:,.0f}",
                "NET DIFF": "{:+.0f}",
                "W/C": "{}"}))
