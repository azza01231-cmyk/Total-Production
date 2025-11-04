# app.py
import streamlit as st
import pandas as pd

st.set_page_config(page_title="NORPETCO Production Summary", layout="wide")
st.title("NORPETCO Daily Production Summary")

uploaded_file = st.file_uploader("Upload Excel Report (.xlsm or .xlsx)", type=["xlsm", "xlsx"])

if uploaded_file:
    try:
        # Step 1: Find the header row that contains "RUNNING WELLS"
        df_raw = pd.read_excel(uploaded_file, sheet_name="Report", header=None)

        header_row = None
        for i, row in df_raw.iterrows():
            if row.astype(str).str.contains("RUNNING WELLS", case=False, na=False).any():
                header_row = i
                break

        if header_row is None:
            st.error("Could not find 'RUNNING WELLS' in the sheet.")
            st.stop()

        # Step 2: Read the actual data table using that row as header
        df = pd.read_excel(
            uploaded_file,
            sheet_name="Report",
            header=header_row,
            skiprows=header_row + 1  # skip the header row itself
        )

        # Clean column names
        df.columns = df.columns.astype(str).str.strip()

        # Step 3: Map exact columns from your table
        col_map = {
            "RUNNING WELLS": "Well Name",
            "Field": "Production Zone",  # This is the zone (Ferdaus, Bahariya, etc.)
            "TOTAL PRODUCTION.3": "TOTAL PRODUCTION",  # Net BO
            "TOTAL PRODUCTION.4": "NET DIFF",          # Net diff. BO
            "W/C.1": "W/C"                             # W/C %
        }

        # Verify all required columns exist
        missing = [c for c in col_map.keys() if c not in df.columns]
        if missing:
            st.error(f"Missing columns: {missing}")
            st.write("Detected columns:", list(df.columns))
            st.stop()

        # Select and rename
        wells_df = df[list(col_map.keys())].copy()
        wells_df.rename(columns=col_map, inplace=True)

        # Step 4: Clean data
        wells_df["TOTAL PRODUCTION"] = pd.to_numeric(wells_df["TOTAL PRODUCTION"], errors="coerce")
        wells_df["NET DIFF"] = pd.to_numeric(wells_df["NET DIFF"], errors="coerce")

        # Remove rows without well name or production
        wells_df = wells_df.dropna(subset=["Well Name", "TOTAL PRODUCTION"], how="any")

        # Step 5: Filter only wells with Net diff. BO
        wells_df = wells_df[wells_df["NET DIFF"].notna()]

        # Step 6: Compute totals
        total_prod = wells_df["TOTAL PRODUCTION"].sum()
        total_diff = wells_df["NET DIFF"].sum()

        # Step 7: Add TOTAL row
        total_row = pd.DataFrame({
            "Well Name": ["TOTAL"],
            "Production Zone": [""],
            "TOTAL PRODUCTION": [total_prod],
            "NET DIFF": [total_diff],
            "W/C": [""]
        })

        final_df = pd.concat([wells_df[["Well Name", "Production Zone", "TOTAL PRODUCTION", "NET DIFF", "W/C"]], total_row], ignore_index=True)

        # Step 8: Display
        st.dataframe(
            final_df.style.format({
                "TOTAL PRODUCTION": "{:,.0f}",
                "NET DIFF": "{:+,.0f}",
                "W/C": "{}"
            }).apply(lambda x: ['font-weight: bold' if x.name == len(final_df)-1 else '' for _ in x], axis=1),
            use_container_width=True
        )

        st.success("Table generated successfully!")

    except Exception as e:
        st.error(f"Error: {e}")
        st.write("Tip: Make sure the sheet name is exactly 'Report' and the table starts with 'RUNNING WELLS'.")

else:
    st.info("Please upload your NORPETCO production report.")
