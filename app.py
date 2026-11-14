
import streamlit as st
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import pandas as pd
import re

st.set_page_config(
    page_title="MEP BOQ Automation",
    page_icon="📊",
    layout="wide"
)

st.title("MEP BOQ Automation System")
st.subheader("Client BOQ to Separate System-wise Excel Files")

SYSTEM_KEYWORDS = {
    "Fire Alarm System": [
        "fire alarm", "smoke detector", "heat detector",
        "mcp", "manual call point", "hooter", "sounder",
        "fire panel", "beam detector", "monitor module",
        "control module", "telephone jack", "strobe"
    ],
    "Public Address System": [
        "public address", "pa system", "speaker",
        "amplifier", "microphone", "paviro",
        "voice evacuation", "ceiling speaker"
    ],
    "CCTV System": [
        "cctv", "camera", "nvr", "dvr",
        "video surveillance", "ip camera"
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
        "dimming", "dali controller"
    ],
    "KNX Automation": [
        "knx", "knx gateway", "knx switch",
        "knx sensor", "knx actuator"
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

SYSTEMS = list(SYSTEM_KEYWORDS.keys()) + [
    "Other ELV Systems",
    "Unassigned"
]


def identify_system(description, part_code=""):
    text = (
        str(description) + " " + str(part_code)
    ).lower()

    matches = []

    for system, keywords in SYSTEM_KEYWORDS.items():
        for keyword in keywords:
            if keyword in text:
                matches.append(system)
                break

    if len(matches) == 1:
        return matches[0]

    # Ambiguous or unidentified items require manual review
    return "Unassigned"


def safe_filename(name):
    return re.sub(
        r"[^A-Za-z0-9_-]+",
        "_",
        name
    )


def create_system_excel(original_bytes, sheet_name, header_row, selected_rows):
    wb = load_workbook(BytesIO(original_bytes))
    ws = wb[sheet_name]

    # Remove all unrelated rows, from bottom to top
    for row_num in range(ws.max_row, header_row, -1):
        if row_num not in selected_rows:
            ws.delete_rows(row_num, 1)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output.getvalue()


st.header("1. Upload Client Blank BOQ")

uploaded = st.file_uploader(
    "Upload client Excel file",
    type=["xlsx"],
    help="Upload the original .xlsx BOQ file."
)

if uploaded:

    try:
        original_bytes = uploaded.getvalue()

        workbook = load_workbook(
            BytesIO(original_bytes),
            data_only=False
        )

        sheet_name = st.selectbox(
            "Select worksheet",
            workbook.sheetnames
        )

        ws = workbook[sheet_name]

        st.success("Excel file uploaded successfully!")

        st.write("Total rows:", ws.max_row)
        st.write("Total columns:", ws.max_column)

        st.header("2. Select Header Row")

        header_row = st.number_input(
            "Which row contains the column headings?",
            min_value=1,
            max_value=max(1, ws.max_row),
            value=1,
            step=1
        )

        header_row = int(header_row)

        # Create unique column labels using actual Excel positions
        column_options = {}

        for col_num in range(1, ws.max_column + 1):
            value = ws.cell(
                row=header_row,
                column=col_num
            ).value

            if value is None or str(value).strip() == "":
                value = "Unnamed"

            letter = get_column_letter(col_num)

            label = f"{letter} - {str(value).strip()}"

            if label in column_options:
                label += f" (Column {col_num})"

            column_options[label] = col_num

        if not column_options:
            st.error("No columns found in selected header row.")
            st.stop()

        st.header("3. Select BOQ Columns")

        description_label = st.selectbox(
            "Select Description / Item Details column",
            list(column_options.keys())
        )

        code_label = st.selectbox(
            "Select Part Code column (if available)",
            ["None"] + list(column_options.keys())
        )

        description_col = column_options[description_label]

        if code_label == "None":
            code_col = None
        else:
            code_col = column_options[code_label]

        if st.button(
            "Identify Different Systems",
            type="primary"
        ):

            records = []

            # Use actual Excel row numbers, not DataFrame indexes
            for row_num in range(header_row + 1, ws.max_row + 1):

                values = [
                    ws.cell(row=row_num, column=col).value
                    for col in range(1, ws.max_column + 1)
                ]

                # Skip completely empty rows
                if all(v is None for v in values):
                    continue

                description = ws.cell(
                    row=row_num,
                    column=description_col
                ).value

                part_code = (
                    ws.cell(row=row_num, column=code_col).value
                    if code_col else ""
                )

                suggestion = identify_system(
                    description,
                    part_code
                )

                records.append({
                    "Excel Row": row_num,
                    "Description": str(description or ""),
                    "Suggested System": suggestion,
                    "Final System": suggestion
                })

            if not records:
                st.warning("No BOQ items found below the header.")
            else:
                st.session_state["boq_records"] = pd.DataFrame(records)
                st.session_state["source_bytes"] = original_bytes
                st.session_state["source_sheet"] = sheet_name
                st.session_state["source_header"] = header_row
                st.session_state["source_filename"] = uploaded.name
                st.session_state.pop("generated_files", None)

                st.success(
                    f"{len(records)} BOQ rows identified."
                )

        if "boq_records" in st.session_state:

            st.header("4. Review System Identification")

            st.info(
                "Check every suggestion. You can manually change "
                "the Final System before generating files."
            )

            edited = st.data_editor(
                st.session_state["boq_records"],
                use_container_width=True,
                hide_index=True,
                disabled=[
                    "Excel Row",
                    "Description",
                    "Suggested System"
                ],
                column_config={
                    "Final System": st.column_config.SelectboxColumn(
                        "Final System",
                        options=SYSTEMS,
                        required=True
                    )
                },
                key="review_editor"
            )

            st.session_state["boq_records"] = edited

            st.write(
                "Total items:",
                len(edited)
            )

            st.write(
                "Unassigned items:",
                len(
                    edited[
                        edited["Final System"] == "Unassigned"
                    ]
                )
            )

            st.header("5. Generate Separate Excel BOQs")

            if st.button(
                "Generate System-wise BOQs",
                type="primary"
            ):

                generated_files = []

                zip_buffer = BytesIO()

                systems_found = edited["Final System"].unique()

                try:
                    with ZipFile(
                        zip_buffer,
                        "w",
                        ZIP_DEFLATED
                    ) as zip_file:

                        for system in systems_found:

                            if system == "Unassigned":
                                continue

                            selected_rows = set(
                                edited[
                                    edited["Final System"] == system
                                ]["Excel Row"].tolist()
                            )

                            if not selected_rows:
                                continue

                            file_content = create_system_excel(
                                st.session_state["source_bytes"],
                                st.session_state["source_sheet"],
                                st.session_state["source_header"],
                                selected_rows
                            )

                            filename = (
                                safe_filename(system)
                                + "_BOQ.xlsx"
                            )

                            zip_file.writestr(
                                filename,
                                file_content
                            )

                            generated_files.append({
                                "system": system,
                                "filename": filename,
                                "content": file_content
                            })

                    st.session_state["generated_files"] = generated_files
                    st.session_state["generated_zip"] = zip_buffer.getvalue()

                except Exception as error:
                    st.error(
                        f"Excel generation error: {error}"
                    )

            if "generated_files" in st.session_state:

                files = st.session_state["generated_files"]

                st.success(
                    f"{len(files)} separate system files generated!"
                )

                for item in files:

                    st.download_button(
                        label="Download " + item["system"],
                        data=item["content"],
                        file_name=item["filename"],
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_" + item["filename"]
                    )

                st.download_button(
                    label="Download ALL System BOQs (ZIP)",
                    data=st.session_state["generated_zip"],
                    file_name="MEP_System_Wise_BOQs.zip",
                    mime="application/zip",
                    type="primary"
                )

    except Exception as error:
        st.error(
            f"Unable to process Excel: {error}"
        )

st.divider()

st.caption(
    "MEP BOQ Automation | System Separation Version 2"
)
