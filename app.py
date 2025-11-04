# --------------------------------------------------------------
# app.py   –   Streamlit Production-Report Summary
# --------------------------------------------------------------
import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="NORPETCO Production Summary", layout="wide")
st.title("NORPETCO Production Report Summary")

uploaded_file = st.file_uploader(
    "Upload the Excel Report (.xlsm or .xlsx)",
    type=["xlsm", "xlsx"],
)

# ------------------------------------------------------------------
# Helper: find the first row that looks like the report header
# ------------------------------------------------------------------
def find_header_end(df_raw: pd.DataFrame) -> int:
    """
    The header ends with the row that contains the word “Date:” (or “From:”).
    Everything **above** that row belongs to the logo / address block.
    """
    for i, row in df_raw.iterrows():
        txt = " ".join(row.astype(str).str.lower())
        if "date:" in txt or "from:" in txt:
            return i + 1                     # data starts on the next line
    return 0                                 # fallback – start from top


# ------------------------------------------------------------------
# Main processing
# ------------------------------------------------------------------
if uploaded_file:
    try:
        # 1. Read the whole sheet without any header assumption
        raw = pd.read_excel(uploaded_file, sheet_name="Report", header=None)

        # 2. Skip the logo / address block
        data_start = find_header_end(raw)

        # 3. Re-read from the first real data line (use the line that contains
        #    “RUNNING WELLS” as the column-header line)
        for i in range(data_start, len(raw)):
            if raw.iloc[i].astype(str).str.contains("RUNNING WELLS", case=False, na=False).any():
                header_row = i
                break
        else:
            st.error("Could not locate the 'RUNNING WELLS' row.")
            st.stop()

        df = pd.read_excel(
            uploaded_file,
            sheet_name="Report",
            header=header_row,          # column names are on this line
            skiprows=header_row + 1,    # skip the header itself when reading data
        )

        # ------------------------------------------------------------------
        # 4. Clean column names
        # ------------------------------------------------------------------
        df.columns = df.columns.astype(str).str.strip()

        # ------------------------------------------------------------------
        # 5. Auto-detect the columns we need
        # ------------------------------------------------------------------
        well_col = next((c for c in df.columns if "well" in c.lower()), None)
        net_bo_col = next(
            (c for c in df.columns if "net" in c.lower() and "bo" in c.lower() and "diff" not in c.lower()),
            None,
        )
        net_diff_col = next(
            (c for c in df.columns if "net" in c.lower() and "diff" in c.lower()), None
        )
        wc_col = next((c for c in df.columns if "w/c" in c.lower() or "wc" in c.lower() or "%" in c), None)

        st.write(
            "**Detected columns**",
            {
                "Well / Zone column": well_col,
                "Net BO column": net_bo_col,
                "Net diff. BO column": net_diff_col,
                "W/C column": wc_col,
            },
        )

        if not all([well_col, net_bo_col, net_diff_col, wc_col]):
            st.error("One or more required columns could not be found.")
            st.stop()

        # ------------------------------------------------------------------
        # 6. Build the final dataframe
        # ------------------------------------------------------------------
        # Start with a copy that contains *all* rows (including zone headers)
        work = df[[well_col, net_bo_col, net_diff_col, wc_col]].copy()
        work["Production Zone"] = ""

        current_zone = ""
        zone_rows = []

        for idx, row in work.iterrows():
            # ---- zone header detection (empty numeric cells) ----
            if pd.isna(row[net_bo_col]) and pd.isna(row[net_diff_col]) and pd.notna(row[well_col]):
                current_zone = str(row[well_col]).strip()
                zone_rows.append(idx)
                continue

            # ---- normal well row ----
            if pd.notna(row[net_bo_col]) or pd.notna(row[net_diff_col]):
                work.at[idx, "Production Zone"] = current_zone

        # Remove the pure zone-header rows (they have no production numbers)
        work = work.drop(index=zone_rows, errors="ignore")

        # Rename for the final table
        work.rename(
            columns={
                well_col: "Well Name",
                net_bo_col: "TOTAL PRODUCTION",
                net_diff_col: "NET DIFF",
                wc_col: "W/C",
            },
            inplace=True,
        )

        # ------------------------------------------------------------------
        # 7. Keep only wells that actually have a Net diff. BO value
        # ------------------------------------------------------------------
        work["TOTAL PRODUCTION"] = pd.to_numeric(work["TOTAL PRODUCTION"], errors="coerce")
        work["NET DIFF"] = pd.to_numeric(work["NET DIFF"], errors="coerce")

        work = work[work["NET DIFF"].notna()].copy()

        # ------------------------------------------------------------------
        # 8. Totals row
        # ------------------------------------------------------------------
        total_prod = work["TOTAL PRODUCTION"].sum()
        total_diff = work["NET DIFF"].sum()

        total_row = pd.DataFrame(
            {
                "Well Name": ["TOTAL"],
                "Production Zone": [""],
                "TOTAL PRODUCTION": [total_prod],
                "NET DIFF": [total_diff],
                "W/C": [""],
            }
        )

        final_df = pd.concat([work[["Well Name", "Production Zone", "TOTAL PRODUCTION", "NET DIFF", "W/C"]], total_row],
                             ignore_index=True)

        # ------------------------------------------------------------------
        # 9. Display
        # ------------------------------------------------------------------
        st.dataframe(
            final_df.style.format(
                {
                    "TOTAL PRODUCTION": "{:,.0f}",
                    "NET DIFF": "{:+.0f}",
                    "W/C": "{}",
                }
            ),
            use_container_width=True,
        )

        st.success("Table generated successfully!")

    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Please upload the NORPETCO production report to continue.")
