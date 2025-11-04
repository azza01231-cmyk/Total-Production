import streamlit as st
import pandas as pd

st.set_page_config(page_title="Production Report Summary", layout="wide")

st.title("🛢️ Production Report Summary")

# Upload Excel file
uploaded_file = st.file_uploader("Upload the Excel Report (.xlsm)", type=["xlsm", "xlsx"])

if uploaded_file:
    try:
        # Read 'Report' sheet
        df = pd.read_excel(uploaded_file, sheet_name="Report")

        # Clean columns (remove spaces, etc.)
        df.columns = df.columns.str.strip()

        # Filter wells that have Net Diff. BO in Total Production
        wells_df = df[[
            "Well Name",
            "TOTAL PRODUCTION",
            "PRODUCTION Zone",
            "NET DIFF",
            "W/C"
        ]].dropna(subset=["TOTAL PRODUCTION"])

        # Convert numeric columns
        wells_df["TOTAL PRODUCTION"] = pd.to_numeric(wells_df["TOTAL PRODUCTION"], errors="coerce")
        wells_df["NET DIFF"] = pd.to_numeric(wells_df["NET DIFF"], errors="coerce")

        # Compute totals
        total_production = wells_df["TOTAL PRODUCTION"].sum()
        total_net_diff = wells_df["NET DIFF"].sum()

        # Add totals row
        total_row = pd.DataFrame({
            "Well Name": ["TOTAL"],
            "TOTAL PRODUCTION": [total_production],
            "PRODUCTION Zone": [""],
            "NET DIFF": [total_net_diff],
            "W/C": [""]
        })

        wells_df = pd.concat([wells_df, total_row], ignore_index=True)

        # Display the table
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
        st.error(f"Error reading Excel file: {e}")
else:
    st.info("📂 Please upload the production Excel file to continue.")
