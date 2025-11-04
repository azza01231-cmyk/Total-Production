# app.py - Streamlit app to extract and summarize production report
import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="NORPETCO Production Summary", layout="wide")
st.title("NORPETCO Daily Production Summary")

uploaded_file = st.file_uploader("Upload Excel Report (.xlsm or .xlsx)", type=["xlsm", "xlsx"])

if uploaded_file:
    try:
        # Step 1: Read raw to find the starting row of the table header
        df_raw = pd.read_excel(uploaded_file, sheet_name="Report", header=None)

        header_first_row = None
        for i, row in df_raw.iterrows():
            if row.astype(str).str.contains("RUNNING WELLS", case=False, na=False).any():
                header_first_row = i
                break

        if header_first_row is None:
            st.error("Could not find the table header starting with 'RUNNING WELLS'.")
            st.stop()

        # Step 2: The header spans two rows: first with main headers, second with sub-headers
        # Read the Excel with multi-index header
        df = pd.read_excel(
            uploaded_file,
            sheet_name="Report",
            header=[header_first_row, header_first_row + 1]
        )

        # Step 3: Define the multi-index column names based on the provided table structure
        zone_col = ('Field', np.nan)  # Production Zone (e.g., Ferdaus, Sidra)
        well_col = ('RUNNING WELLS', np.nan)  # Well Name
        net_bo_col = ('TOTAL PRODUCTION', 'Net\nBO')  # Total Production (Net BO)
        net_diff_col = ('TOTAL PRODUCTION', 'Net diff. BO')  # Net Diff BO
        wc_col = ('W/C', '%')  # W/C %

        # Verify all required columns exist
        missing_cols = [col for col in [zone_col, well_col, net_bo_col, net_diff_col, wc_col] if col not in df.columns]
        if missing_cols:
            st.error(f"Missing expected columns in the table: {missing_cols}")
            st.write("Detected columns:", list(df.columns))
            st.stop()

        # Step 4: Select relevant columns and rename for simplicity
        wells_df = df[[zone_col, well_col, net_bo_col, net_diff_col, wc_col]].copy()
        wells_df.columns = ["Production Zone", "Well Name", "TOTAL PRODUCTION", "NET DIFF", "W/C"]

        # Step 5: Forward-fill the Production Zone (since it's merged/grouped in Excel)
        wells_df["Production Zone"] = wells_df["Production Zone"].ffill()

        # Step 6: Clean data - drop rows without well names or production values
        wells_df = wells_df.dropna(subset=["Well Name", "TOTAL PRODUCTION"], how="any")

        # Exclude any footer rows like "TOTAL" or "CUM. PROD." if present
        wells_df = wells_df[~wells_df["Well Name"].astype(str).str.contains("TOTAL|CUM", case=False, na=False)]

        # Step 7: Convert to numeric, handling errors like #VALUE!
        wells_df["TOTAL PRODUCTION"] = pd.to_numeric(wells_df["TOTAL PRODUCTION"], errors="coerce")
        wells_df["NET DIFF"] = pd.to_numeric(wells_df["NET DIFF"], errors="coerce")

        # Step 8: Filter only wells that have a Net Diff BO value (as per your request)
        wells_df = wells_df[wells_df["NET DIFF"].notna()]

        # Step 9: Compute totals
        total_prod = wells_df["TOTAL PRODUCTION"].sum()
        total_diff = wells_df["NET DIFF"].sum()

        # Step 10: Add TOTAL row
        total_row = pd.DataFrame({
            "Well Name": ["TOTAL"],
            "Production Zone": [""],
            "TOTAL PRODUCTION": [total_prod],
            "NET DIFF": [total_diff],
            "W/C": [""]
        })

        final_df = pd.concat([wells_df, total_row], ignore_index=True)

        # Step 11: Display the styled table
        st.dataframe(
            final_df.style.format({
                "TOTAL PRODUCTION": "{:,.0f}",
                "NET DIFF": "{:+,.0f}",
                "W/C": "{}"
            }).apply(lambda x: ['font-weight: bold' if x.name == len(final_df)-1 else '' for _ in x], axis=1),
            use_container_width=True
        )

        st.success("Table extracted and summarized successfully!")

    except Exception as e:
        st.error(f"Error processing the Excel file: {e}")
        st.write("Tip: Ensure the sheet is named 'Report' and matches the provided table structure.")

else:
    st.info("Please upload your NORPETCO production report.")
