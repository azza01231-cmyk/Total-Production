import streamlit as st
import pandas as pd

st.set_page_config(page_title="Production Report Summary", layout="wide")

st.title("🛢️ Production Report Summary")

uploaded_file = st.file_uploader("Upload the Excel Report (.xlsm)", type=["xlsm", "xlsx"])

if uploaded_file:
    try:
        # Read the 'Report' sheet
        df = pd.read_excel(uploaded_file, sheet_name="Report")

        # Clean up column names (strip spaces)
        df.columns = df.columns.str.strip()

        # Print columns for debugging
        st.write("Detected columns:", list(df.columns))

        # Select relevant columns from your provided image
        wells_df = df[[
            "RUNNING WELLS",
            "TOTAL PRODUCTION.1",   # Net BO
            "TOTAL PRODUCTION.2",   # Net diff. BO
            "W/C",                  # W/C %
        ]].copy()

        # Rename for clarity
        wells_df.columns = ["Well Name", "TOTAL PRODUCTION", "NET DIFF", "W/C"]

        # Convert numeric columns
        wells_df["TOTAL PRODUCTION"] = pd.to_numeric(wells_df["TOTAL PRODUCTION"], errors="coerce")
        wells_df["NET DIFF"] = pd.to_numeric(wells_df["NET DIFF"], errors="coerce")

        # Compute totals
        total_production = wells_df["TOTAL PRODUCTION"].sum()
        total_net_diff = wells_df["NET DIFF"].sum()

        # Add total row
        total_row = pd.DataFrame({
            "Well Name": ["TOTAL"],
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
                "W/C": "{}"
            }),
            use_container_width=True
        )

        st.success("✅ Table generated successfully!")

    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
else:
    st.info("📂 Please upload the production Excel file to continue.")
