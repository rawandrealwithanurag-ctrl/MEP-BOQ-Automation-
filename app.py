
import streamlit as st
import pandas as pd
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from openpyxl import load_workbook
import re

st.set_page_config(
    page_title="MEP BOQ Automation",
    page_icon="📊",
    layout="wide"
)

st.title("MEP BOQ Automation")
st.subheader("Client BOQ - System Wise Separation")

SYSTEM_KEYWORDS = {
    "Fire Alarm System": [
        "fire alarm", "smoke detector", "heat detector",
        "mcp", "manual call point", "hooter", "sounder",
        "fire panel", "beam detector", "monitor module",
        "control module", "telephone jack"
    ],
    "Public Address System": [
        "public address", "pa system", "speaker",
        "amplifier", "microphone", "paviro",
        "voice evacuation"
    ],
    "CCTV System": [
        "cctv", "camera", "nvr", "dvr",
        "video surveillance"
    ],
    "Access Control System": [
        "access control", "card reader", "door controller",
        "magnetic lock", "em lock", "access reader"
    ],
    "BMS System": [
        "bms", "ddc", "bacnet", "modbus",
        "field controller", "bms panel"
    ],
    "DALI Automation": [
        "dali", "dali gateway", "dali driver",
        "dimming"
    ],
    "KNX Automation": [
        "knx", "knx gateway", "knx switch",
        "knx sensor"
    ],
    "Structured Cabling": [
        "structured cabling", "patch panel",
        "cat6", "cat 6", "fiber optic",
        "network rack", "data point"
    ],
    "Fire Fighting System": [
        "sprinkler", "fire hydrant", "fire pump",
        "landing valve", "hose reel",
        "fire extinguisher"
    ],
    "Plumbing System": [
        "plumbing", "water supply", "drainage",
        "soil pipe", "waste pipe", "sanitary"
    ],
    "HVAC System": [
        "hvac", "ahu", "fcu", "chiller",
        "ducting", "diffuser", "grille"
    ]
}

SYSTEMS = list(SYSTEM_KEYWORDS.keys()) + ["Other ELV Systems", "Unassigned"]


def identify_system(description, part_code=""):
    text = (str(description) + " " + str(part_code)).lower()

    matches = []

    for system, keywords in SYSTEM_KEYWORDS.items():
        if any(word in text for word in keywords):
            matches.append(system)

    if len(matches) == 1:
        return matches[0]

    return "Unassigned"


def excel_bytes(wb):
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


st.header("1. Upload Client Blank BOQ")

uploaded = st.file_uploader(
    "Upload original Excel file (.xlsx)",
    type=["xlsx"]
)

if uploaded:

    try:
        file_content = uploaded.getvalue()

        wb_original = load_workbook(
            BytesIO(file_content)
        )

        sheet_name = st.selectbox(
            "Select worksheet",
            wb_original.sheetnames
        )

        ws = wb_original[sheet_name]

        # Read worksheet values
        data = list(ws.values)

        if not data:
            st.error("Excel sheet is empty.")
            st.stop()

        header_row = st.number_input(
            "Header row number",
            min_value=1,
            max_value=ws.max_row,
            value=1
        )

        headers = list(
            ws.iter_rows(
                min_row=int(header_row),
                max_row=int(header_row),
                values_only=True
            )
        )[0]

        headers = [
            str(h) if h is not None else f"Column {i+1}"
            for i, h in enumerate(headers)
        ]

        df = pd.read_excel(
            BytesIO(file_content),
            sheet_name=sheet_name,
            header=int(header_row)-1,
            engine="openpyxl"
        )

        st.success("Client BOQ uploaded successfully!")

        st.write("Total Excel rows:", ws.max_row)
        st.write("Total BOQ records:", len(df))

        st.header("2. Select BOQ Columns")

        description_col = st.selectbox(
            "Select Item Description column",
            headers
        )

        code_col = st.selectbox(
            "Select Part Code column (if available)",
            ["None"] + headers
        )

        if st.button("Identify Different Systems", type="primary"):

            suggestions = []

            for _, row in df.iterrows():

                description = row.get(description_col, "")

                part_code = (
                    row.get(code_col, "")
                    if code_col != "None"
                    else ""
                )

                suggestions.append(
                    identify_system(description, part_code)
                )

            st.session_state["boq_df"] = df.copy()
            st.session_state["suggestions"] = suggestions
            st.session_state["excel_content"] = file_content
            st.session_state["selected_sheet"] = sheet_name
            st.session_state["header_row"] = int(header_row)

        if "boq_df" in st.session_state:

            st.header("3. Review System Identification")

            original_df = st.session_state["boq_df"]

            review_df = pd.DataFrame({
                "Excel Row": [
                    i + st.session_state["header_row"] + 1
                    for i in range(len(original_df))
                ],
                "Description": original_df[description_col].astype(str),
                "Suggested System": st.session_state["suggestions"],
                "Final System": st.session_state["suggestions"]
            })

            st.info(
                "Review the suggested systems. Change incorrect "
                "assignments manually before downloading."
            )

            edited_df = st.data_editor(
                review_df,
                use_container_width=True,
                hide_index=True,
                disabled=["Excel Row", "Description", "Suggested System"],
                column_config={
                    "Final System": st.column_config.SelectboxColumn(
                        "Final System",
                        options=SYSTEMS,
                        required=True
                    )
                },
                key="system_review"
            )

            st.header("4. Generate Separate Excel Files")

            if st.button("Generate System-wise BOQs", type="primary"):

                grouped_systems = edited_df["Final System"].unique()

                zip_buffer = BytesIO()
                generated = []

                with ZipFile(
                    zip_buffer, "w", ZIP_DEFLATED
                ) as zip_file:

                    for system in grouped_systems:

                        if system == "Unassigned":
                            continue

                        selected_rows = edited_df[
                            edited_df["Final System"] == system
                        ]["Excel Row"].tolist()

                        # Load original workbook to preserve formatting
                        wb = load_workbook(
                            BytesIO(st.session_state["excel_content"])
                        )

                        ws = wb[st.session_state["selected_sheet"]]

                        keep_rows = set(selected_rows)

                        # Remove rows not belonging to this system
                        # Process from bottom to top
                        for r in range(ws.max_row, 0, -1):

                            if r <= st.session_state["header_row"]:
                                continue

                            if r not in keep_rows:
                                ws.delete_rows(r, 1)

                        output = excel_bytes(wb)

                        safe_name = re.sub(
                            r"[^A-Za-z0-9]+",
                            "_",
                            system
                        )

                        filename = safe_name + "_BOQ.xlsx"

                        zip_file.writestr(
                            filename,
                            output.getvalue()
                        )

                        generated.append(
                            (system, filename, output.getvalue())
                        )

                st.session_state["generated"] = generated
                st.session_state["zip_data"] = zip_buffer.getvalue()

            if "generated" in st.session_state:

                st.success(
                    f"{len(st.session_state['generated'])} system files generated!"
                )

                for system, filename, content in st.session_state["generated"]:

                    st.download_button(
                        label=f"Download {system} Excel",
                        data=content,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_" + filename
                    )

                st.download_button(
                    "Download ALL Systems (ZIP)",
                    data=st.session_state["zip_data"],
                    file_name="MEP_All_System_BOQs.zip",
                    mime="application/zip",
                    type="primary"
                )

            unassigned = edited_df[
                edited_df["Final System"] == "Unassigned"
            ]

            if len(unassigned) > 0:
                st.warning(
                    f"{len(unassigned)} rows are unassigned. "
                    "Review these rows manually."
                )

    except Exception as e:
        st.error(f"Error processing Excel: {e}")
        st.info(
            "Check the Excel file and selected worksheet. "
            "For complex merged-cell templates, custom handling may be needed."
        )

st.divider()
st.caption("MEP BOQ Automation | System Separation Version 1")
