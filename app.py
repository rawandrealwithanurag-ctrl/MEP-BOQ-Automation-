
import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
import re

st.set_page_config(page_title="MEP BOQ Automation", page_icon="📊", layout="wide")

st.title("MEP BOQ Automation System")
st.caption("Client BOQ to OEM Rated BOQ")

SYSTEMS = [
    "Fire Alarm System",
    "Public Address System",
    "CCTV System",
    "Access Control System",
    "BMS System",
    "DALI Automation",
    "KNX Automation",
    "Structured Cabling",
    "Fire Fighting System",
    "Plumbing System",
    "HVAC System",
    "Other ELV Systems",
    "Unassigned"
]

KEYWORDS = {
    "Fire Alarm System": ["fire alarm", "smoke detector", "heat detector", "mcp", "hooter", "sounder", "fire panel", "beam detector", "monitor module", "control module"],
    "Public Address System": ["public address", "pa system", "speaker", "amplifier", "microphone", "paviro", "voice evacuation"],
    "CCTV System": ["cctv", "camera", "nvr", "dvr", "video surveillance"],
    "Access Control System": ["access control", "reader", "card reader", "door controller", "magnetic lock", "em lock"],
    "BMS System": ["bms", "ddc", "bacnet", "modbus", "field controller", "sensor", "actuator"],
    "DALI Automation": ["dali", "dali gateway", "dali driver", "dimming"],
    "KNX Automation": ["knx", "knx gateway", "knx switch", "knx sensor"],
    "Structured Cabling": ["structured cabling", "data point", "patch panel", "network rack", "cat6", "cat 6", "fiber optic"],
    "Fire Fighting System": ["sprinkler", "fire hydrant", "fire pump", "landing valve", "hose reel", "fire extinguisher"],
    "Plumbing System": ["plumbing", "water supply", "drainage", "soil pipe", "waste pipe", "sanitary"],
    "HVAC System": ["hvac", "ahu", "fcu", "chiller", "ducting", "diffuser", "grille"]
}

def identify_system(value):
    text = str(value).lower()
    matches = []
    for system, words in KEYWORDS.items():
        if any(word in text for word in words):
            matches.append(system)
    if len(matches) == 1:
        return matches[0]
    return "Unassigned"

def make_excel(data):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        data.to_excel(writer, index=False, sheet_name="BOQ")
    output.seek(0)
    return output

st.header("1. Upload Client Blank BOQ")

file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if file:
    try:
        excel = pd.ExcelFile(file)
        sheet = st.selectbox("Select BOQ Sheet", excel.sheet_names)
        df = pd.read_excel(file, sheet_name=sheet)

        st.success("Client BOQ uploaded successfully!")
        st.write("Total rows:", len(df))
        st.write("Total columns:", len(df.columns))

        st.header("2. Identify BOQ Systems")

        columns = list(df.columns)
        description_col = st.selectbox(
            "Select Description / Item Details Column",
            columns
        )

        code_col = st.selectbox(
            "Select Part Code Column (optional)",
            ["None"] + columns
        )

        if st.button("Identify Systems", type="primary"):
            combined = df[description_col].fillna("").astype(str)

            if code_col != "None":
                combined = combined + " " + df[code_col].fillna("").astype(str)

            df["Suggested System"] = combined.apply(identify_system)
            df["Final System"] = df["Suggested System"]
            st.session_state["boq_data"] = df
            st.session_state["boq_sheet"] = sheet

        if "boq_data" in st.session_state:
            st.header("3. Review and Correct Systems")
            st.info("Check the suggested systems. You can manually change any incorrect assignment.")

            edited = st.data_editor(
                st.session_state["boq_data"],
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                column_config={
                    "Suggested System": st.column_config.TextColumn(
                        disabled=True
                    ),
                    "Final System": st.column_config.SelectboxColumn(
                        options=SYSTEMS,
                        required=True
                    )
                },
                key="system_editor"
            )

            st.session_state["boq_data"] = edited

            st.header("4. Download Separate System BOQs")

            available = [
                s for s in SYSTEMS
                if s in edited["Final System"].values
                and s != "Unassigned"
            ]

            if available:
                for system in available:
                    system_data = edited[
                        edited["Final System"] == system
                    ].drop(columns=["Suggested System", "Final System"])

                    st.download_button(
                        label=f"Download {system} Excel",
                        data=make_excel(system_data),
                        file_name=re.sub(r"[^A-Za-z0-9]+", "_", system) + ".xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_" + system
                    )

                zip_buffer = BytesIO()

                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:
                    for system in available:
                        system_data = edited[
                            edited["Final System"] == system
                        ].drop(columns=["Suggested System", "Final System"])

                        excel_file = make_excel(system_data)
                        filename = re.sub(r"[^A-Za-z0-9]+", "_", system) + ".xlsx"
                        z.writestr(filename, excel_file.getvalue())

                zip_buffer.seek(0)

                st.download_button(
                    "Download ALL System BOQs (ZIP)",
                    data=zip_buffer,
                    file_name="MEP_System_Wise_BOQs.zip",
                    mime="application/zip",
                    type="primary"
                )
            else:
                st.warning("No systems identified yet. Please review your assignments.")

            unassigned = edited[
                edited["Final System"] == "Unassigned"
            ]

            if len(unassigned) > 0:
                st.warning(f"{len(unassigned)} rows are still unassigned. Please review them.")

    except Exception as e:
        st.error(f"Error reading BOQ: {e}")

st.divider()
st.caption("MEP BOQ Automation | Version 2")
