
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="MEP BOQ Automation",
    page_icon="📊",
    layout="wide"
)

st.title("MEP BOQ Automation System")
st.subheader("Client BOQ to OEM Rated BOQ")

st.write("Welcome! Manage your project BOQs in one place.")

st.header("Step 1: Upload Client Blank BOQ")

file = st.file_uploader(
    "Upload Client Excel BOQ",
    type=["xlsx", "xls"]
)

if file:
    try:
        excel = pd.ExcelFile(file)
        st.success("BOQ uploaded successfully!")

        sheet = st.selectbox(
            "Select BOQ Sheet",
            excel.sheet_names
        )

        df = pd.read_excel(file, sheet_name=sheet)

        st.subheader("BOQ Preview")
        st.dataframe(df, use_container_width=True)

        st.write("Total Rows:", len(df))

        st.info(
            "Next: We will add system-wise separation, "
            "OEM rate mapping and final client BOQ generation."
        )

    except Exception as e:
        st.error(f"Unable to read file: {e}")

st.divider()
st.caption("MEP BOQ Automation | Version 1.0")
