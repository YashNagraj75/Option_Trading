from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st
import xlsxwriter

st.title("Option Trades Analyzer")

uploaded_file = st.file_uploader("Upload your options CSV file", type=["csv"])


def process_dataframe(df):
    # Calculate new columns
    df["PCR OI"] = df["PUTS_OI"] / df["CALLS_OI"]
    df["PCR Volume"] = df["PUTS_VOLUME"] / df["CALLS_VOLUME"]
    df["CPR OI"] = df["CALLS_OI"] / df["PUTS_OI"]
    df["CPR Volume"] = df["CALLS_VOLUME"] / df["PUTS_VOLUME"]
    df["PCR Sum"] = df["PCR OI"] + df["PCR Volume"]
    df["CPR Sum"] = df["CPR OI"] + df["CPR Volume"]

    # Assign type
    def assign_type(row):
        if row["PCR Sum"] > 6 and row["PCR Sum"] < 15:
            return "Good support"
        elif row["CPR Sum"] > 6 and row["CPR Sum"] < 15:
            return "Good resistance"
        elif row["PCR Sum"] > 15:
            return "Very good support"
        elif row["CPR Sum"] > 15:
            return "Very good resistance"
        else:
            return "-"

    df["type"] = df.apply(assign_type, axis=1)
    # Removed sorting logic
    return df


def to_excel_with_highlight(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]
        # Set wider columns for readability
        for i, col in enumerate(df.columns):
            worksheet.set_column(i, i, 18)
        # Highlight STRIKE column in yellow
        strike_col_idx = df.columns.get_loc("STRIKE")
        yellow_fmt = workbook.add_format({"bg_color": "#FFFF00"})
        worksheet.set_column(strike_col_idx, strike_col_idx, 18, yellow_fmt)
        # Highlight top 3 values in PCR OI, CPR OI, PCR Volume, CPR Volume in green
        green_fmt = workbook.add_format({"bg_color": "#C6EFCE"})
        highlight_cols = [
            "PCR OI",
            "CPR OI",
            "PCR Volume",
            "CPR Volume",
        ]
        for col in highlight_cols:
            if col in df.columns:
                col_idx = df.columns.get_loc(col)
                worksheet.conditional_format(
                    1,
                    col_idx,
                    len(df),
                    col_idx,
                    {"type": "top", "value": 3, "format": green_fmt},
                )
    output.seek(0)
    return output


if uploaded_file is not None:
    # Read the correct header row (second row)
    df = pd.read_csv(uploaded_file, header=1)
    # Rename columns for clarity, including CHNG IN OI
    df = df.rename(
        columns={
            "OI": "CALLS_OI",
            "CHNG IN OI": "CALLS_CHNG_IN_OI",
            "VOLUME": "CALLS_VOLUME",
            "STRIKE": "STRIKE",
            "OI.1": "PUTS_OI",
            "CHNG IN OI.1": "PUTS_CHNG_IN_OI",
            "VOLUME.1": "PUTS_VOLUME",
        }
    )
    # Clean and convert relevant columns to numeric
    numeric_cols = [
        "CALLS_OI",
        "PUTS_OI",
        "CALLS_VOLUME",
        "PUTS_VOLUME",
        "CALLS_CHNG_IN_OI",
        "PUTS_CHNG_IN_OI",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(",", "").replace("-", "0")
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            st.warning(f"Column '{col}' not found during cleaning.")

    st.write("Raw Data (before processing & column selection):")
    st.dataframe(df, use_container_width=True)
    try:
        df_processed = process_dataframe(
            df.copy()
        )  # Use a copy to avoid modifying original df

        # Define the desired columns in the specified order
        final_columns = [
            "PUTS_OI",
            "PUTS_VOLUME",
            "PUTS_CHNG_IN_OI",
            "PCR OI",
            "PCR Volume",
            "PCR Sum",
            "STRIKE",
            "CALLS_OI",
            "CALLS_VOLUME",
            "CPR OI",
            "CPR Volume",
            "CPR Sum",
            "CALLS_CHNG_IN_OI",
            "type",
        ]

        # Filter and reorder the DataFrame
        # Check if all desired columns exist before reindexing
        missing_cols = [col for col in final_columns if col not in df_processed.columns]
        if missing_cols:
            st.error(
                f"Error: The following required columns are missing after processing: {missing_cols}"
            )
        else:
            df_final = df_processed[final_columns]

            st.write("Processed Data (Selected Columns):")
            st.dataframe(df_final, use_container_width=True)
            xlsx_data = to_excel_with_highlight(df_final)
            st.download_button(
                label="Download as XLSX",
                data=xlsx_data,
                file_name="option_analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        st.error(f"Error processing file: {e}")
