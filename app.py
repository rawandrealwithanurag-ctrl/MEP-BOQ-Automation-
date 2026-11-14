import streamlit as st
import pandas as pd
import re
import os
import tempfile
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font
from difflib import SequenceMatcher

st.set_page_config(page_title="MEP BOQ Automation", page_icon="ðŸ“Š", layout="wide")
st.title("MEP BOQ Automation System")
st.caption("System-wise BOQ separation + OEM rate mapping + BOQ comparison")

SYSTEM_KEYWORDS = {
    "Fire Alarm System": ["fire alarm", "smoke detector", "heat detector", "mcp", "manual call point", "hooter", "sounder", "fire panel", "beam detector", "monitor module", "control module", "telephone jack", "strobe"],
    "Public Address System": ["public address", "pa system", "speaker", "amplifier", "microphone", "paviro", "voice evacuation", "ceiling speaker"],
    "CCTV System": ["cctv", "camera", "nvr", "dvr", "video surveillance", "ip camera"],
    "Access Control System": ["access control", "card reader", "door controller", "magnetic lock", "em lock", "access reader"],
    "BMS System": ["bms", "ddc", "bacnet", "modbus", "field controller", "bms panel"],
    "DALI Automation": ["dali", "dali gateway", "dali driver", "dimming", "dali controller"],
    "KNX Automation": ["knx", "knx gateway", "knx switch", "knx sensor", "knx actuator"],
    "Structured Cabling": ["structured cabling", "patch panel", "cat6", "cat 6", "fiber optic", "network rack", "data point"],
    "Fire Fighting System": ["sprinkler", "fire hydrant", "fire pump", "landing valve", "hose reel", "fire extinguisher"],
    "Plumbing System": ["plumbing", "water supply", "drainage", "soil pipe", "waste pipe", "sanitary"],
    "HVAC System": ["hvac", "ahu", "fcu", "chiller", "ducting", "diffuser", "grille"]
}
SYSTEMS = list(SYSTEM_KEYWORDS) + ["Other ELV Systems", "Unassigned"]


def clean_text(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def norm(value):
    text = clean_text(value).lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def identify_system(description, part_code=""):
    text = (clean_text(description) + " " + clean_text(part_code)).lower()
    found = [name for name, words in SYSTEM_KEYWORDS.items() if any(w in text for w in words)]
    return found[0] if len(found) == 1 else "Unassigned"


def safe_filename(name):
    return re.sub(r"[^A-Za-z0-9_-]+", "_", str(name))


def get_header_map(ws, header_row):
    result = {}
    for col in range(1, ws.max_column + 1):
        val = clean_text(ws.cell(header_row, col).value)
        if val:
            result[f"{get_column_letter(col)} - {val}"] = col
    return result


def guess_column(labels, words):
    for label, col in labels.items():
        low = label.lower()
        if any(word in low for word in words):
            return col
    return None


def parse_excel_oem(file_bytes, filename):
    wb = load_workbook(BytesIO(file_bytes), data_only=True)
    sheets = []
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        # Find likely header row in first 15 rows
        best_row, best_score = 1, -1
        for idx, row in enumerate(rows[:15], start=1):
            joined = " ".join(clean_text(x).lower() for x in row if x is not None)
            score = sum(k in joined for k in ["description", "item", "supply", "installation", "rate", "make", "part code", "qty"])
            if score > best_score:
                best_row, best_score = idx, score
        header = [clean_text(x) or f"Column {i+1}" for i, x in enumerate(rows[best_row-1])]
        df = pd.DataFrame(rows[best_row:], columns=header)
        df = df.dropna(how="all")
        if not df.empty:
            sheets.append((ws.title, best_row, df))
    return sheets


def ocr_file(file_bytes, filename):
    try:
        import pytesseract
        from PIL import Image
    except Exception as e:
        raise RuntimeError("OCR packages are not installed. Check requirements.txt and packages.txt.") from e
    ext = os.path.splitext(filename.lower())[1]
    texts = []
    if ext == ".pdf":
        try:
            import pdfplumber
            with pdfplumber.open(BytesIO(file_bytes)) as pdf:
                for page in pdf.pages:
                    txt = page.extract_text() or ""
                    if txt.strip():
                        texts.append(txt)
                    else:
                        try:
                            from pdf2image import convert_from_bytes
                            imgs = convert_from_bytes(file_bytes, first_page=page.page_number, last_page=page.page_number, dpi=220)
                            if imgs:
                                texts.append(pytesseract.image_to_string(imgs[0]))
                        except Exception:
                            texts.append("")
        except Exception as e:
            raise RuntimeError(f"PDF read error: {e}")
    else:
        image = Image.open(BytesIO(file_bytes)).convert("RGB")
        texts.append(pytesseract.image_to_string(image))
    return "\n".join(texts)


def numeric_tokens(line):
    # Capture numbers with commas/decimals; ignore obvious item numbering at line start later
    vals = re.findall(r"(?<![A-Za-z])(?:â‚¹\s*)?\d[\d,]*(?:\.\d{1,2})?", line)
    out = []
    for v in vals:
        try:
            out.append(float(v.replace(",", "").replace("â‚¹", "").strip()))
        except Exception:
            pass
    return out


def ocr_lines_to_df(text):
    records = []
    for line in text.splitlines():
        line = clean_text(line)
        if len(line) < 4:
            continue
        nums = numeric_tokens(line)
        # Keep only rows with at least one likely rate; user can edit extracted rows.
        if not nums:
            continue
        description = line
        for n in nums:
            description = re.sub(r"(?<![A-Za-z])â‚¹?\s*" + re.escape(f"{n:g}") + r"(?!\d)", " ", description)
        description = clean_text(description.strip(" |:-"))
        if not description:
            description = line
        supply = nums[-2] if len(nums) >= 2 else nums[-1]
        install = nums[-1] if len(nums) >= 2 else None
        records.append({"Description": description, "Part Code": "", "Supply Rate": supply, "Installation Rate": install, "Make": "", "Source Text": line})
    return pd.DataFrame(records)


def similarity(a, b):
    na, nb = norm(a), norm(b)
    if not na or not nb:
        return 0.0
    if na == nb:
        return 1.0
    return SequenceMatcher(None, na, nb).ratio()


def find_client_items(ws, header_row, desc_col, spec_col, qty_col, unit_col, code_col):
    items = []
    for r in range(header_row + 1, ws.max_row + 1):
        desc = clean_text(ws.cell(r, desc_col).value) if desc_col else ""
        code = clean_text(ws.cell(r, code_col).value) if code_col else ""
        qty = ws.cell(r, qty_col).value if qty_col else None
        unit = clean_text(ws.cell(r, unit_col).value) if unit_col else ""
        spec = clean_text(ws.cell(r, spec_col).value) if spec_col else ""
        # Ignore section headings, totals and blank lines. A usable item needs description + quantity or serial number.
        if not desc or re.search(r"\b(total|summary)\b", desc, re.I):
            continue
        if qty is None and not code and not re.match(r"^\d+(\.\d+)?$", clean_text(ws.cell(r, 1).value)):
            continue
        items.append({"Client Row": r, "Client Description": desc, "Client Specification": spec, "Client Code": code, "Client Qty": qty, "Client Unit": unit})
    return items


def make_comparison(client_items, oem_df):
    oem = oem_df.copy().reset_index(drop=True)
    if "Description" not in oem.columns:
        oem["Description"] = ""
    used = set()
    rows = []
    for item in client_items:
        best_i, best_score = None, 0
        target = item["Client Description"] + " " + item.get("Client Specification", "")
        for i, row in oem.iterrows():
            if i in used:
                continue
            code_score = 0
            ccode = norm(item.get("Client Code", ""))
            ocode = norm(row.get("Part Code", ""))
            if ccode and ocode and ccode == ocode:
                code_score = 1.0
            score = max(code_score, similarity(target, clean_text(row.get("Description", ""))))
            if score > best_score:
                best_i, best_score = i, score
        if best_i is not None and best_score >= 0.52:
            used.add(best_i)
            o = oem.iloc[best_i]
            status = "Matched" if best_score >= 0.82 else "Review"
            rows.append({
                "Client Row": item["Client Row"], "Client Description": item["Client Description"],
                "Client Qty": item["Client Qty"], "Client Unit": item["Client Unit"],
                "OEM Row": best_i + 1, "OEM Description": clean_text(o.get("Description", "")),
                "Match Score": round(best_score * 100, 1), "Status": status,
                "Supply Rate": o.get("Supply Rate", None), "Installation Rate": o.get("Installation Rate", None),
                "Make": clean_text(o.get("Make", "")), "Approved": best_score >= 0.82
            })
        else:
            rows.append({
                "Client Row": item["Client Row"], "Client Description": item["Client Description"],
                "Client Qty": item["Client Qty"], "Client Unit": item["Client Unit"],
                "OEM Row": None, "OEM Description": "", "Match Score": 0, "Status": "Missing in OEM",
                "Supply Rate": None, "Installation Rate": None, "Make": "", "Approved": False
            })
    extras = []
    for i, row in oem.iterrows():
        if i not in used:
            extras.append({"OEM Row": i + 1, "OEM Description": clean_text(row.get("Description", "")), "Supply Rate": row.get("Supply Rate", None), "Installation Rate": row.get("Installation Rate", None), "Make": clean_text(row.get("Make", "")), "Status": "Extra in OEM"})
    return pd.DataFrame(rows), pd.DataFrame(extras)


def style_comparison_sheet(ws):
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for row in ws.iter_rows(min_row=2):
        status = str(row[7].value or "") if len(row) > 7 else ""
        color = "E2F0D9" if status == "Matched" else ("FFF2CC" if status == "Review" else "FCE4D6")
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=color)


tab_split, tab_rates = st.tabs(["1. System-wise BOQ Splitter", "2. OEM Rate Mapping"])

with tab_split:
    st.subheader("Separate client BOQ into system-wise Excel files")
    uploaded = st.file_uploader("Upload client Excel BOQ", type=["xlsx"], key="split_upload")
    if uploaded:
        source_bytes = uploaded.getvalue()
        wb = load_workbook(BytesIO(source_bytes), data_only=False)
        sheet = st.selectbox("Select worksheet", wb.sheetnames, key="split_sheet")
        ws = wb[sheet]
        header_row = int(st.number_input("Header row number", min_value=1, max_value=max(1, ws.max_row), value=2, key="split_header"))
        labels = get_header_map(ws, header_row)
        if labels:
            desc_label = st.selectbox("Description column", list(labels), index=next((i for i,k in enumerate(labels) if "description" in k.lower()), 0), key="split_desc")
            code_label = st.selectbox("Part code column (optional)", ["None"] + list(labels), key="split_code")
            if st.button("Identify systems", key="split_identify", type="primary"):
                records = []
                for r in range(header_row + 1, ws.max_row + 1):
                    desc = ws.cell(r, labels[desc_label]).value
                    code = ws.cell(r, labels[code_label]).value if code_label != "None" else ""
                    if clean_text(desc):
                        records.append({"Excel Row": r, "Description": clean_text(desc), "Suggested System": identify_system(desc, code), "Final System": identify_system(desc, code)})
                st.session_state["split_records"] = pd.DataFrame(records)
                st.session_state["split_source"] = source_bytes
                st.session_state["split_sheet_name"] = sheet
                st.session_state["split_header_row"] = header_row
            if "split_records" in st.session_state:
                edited = st.data_editor(st.session_state["split_records"], use_container_width=True, hide_index=True, disabled=["Excel Row", "Description", "Suggested System"], column_config={"Final System": st.column_config.SelectboxColumn(options=SYSTEMS, required=True)}, key="split_editor")
                st.session_state["split_records"] = edited
                if st.button("Generate system-wise Excel ZIP", key="split_generate"):
                    output = BytesIO()
                    with ZipFile(output, "w", ZIP_DEFLATED) as zf:
                        for system in edited["Final System"].dropna().unique():
                            if system == "Unassigned":
                                continue
                            selected = set(edited.loc[edited["Final System"] == system, "Excel Row"].astype(int))
                            out_wb = load_workbook(BytesIO(st.session_state["split_source"]))
                            out_ws = out_wb[st.session_state["split_sheet_name"]]
                            for rr in range(out_ws.max_row, st.session_state["split_header_row"], -1):
                                if rr not in selected and rr != st.session_state["split_header_row"]:
                                    out_ws.delete_rows(rr, 1)
                            buf = BytesIO(); out_wb.save(buf)
                            zf.writestr(safe_filename(system) + "_BOQ.xlsx", buf.getvalue())
                    st.download_button("Download all system BOQs", output.getvalue(), "MEP_System_Wise_BOQs.zip", "application/zip", type="primary")

with tab_rates:
    st.subheader("OEM quotation to client BOQ rate mapping")
    st.info("Upload the blank client BOQ and the OEM rated quotation. OEM file can be Excel, PDF or photo. OCR results and uncertain matches must be reviewed before rates are applied.")
    client_file = st.file_uploader("A. Upload client blank BOQ (.xlsx)", type=["xlsx"], key="client_rate_file")
    oem_file = st.file_uploader("B. Upload OEM rated file (Excel / PDF / JPG / PNG)", type=["xlsx", "xls", "pdf", "jpg", "jpeg", "png"], key="oem_rate_file")
    if client_file and oem_file:
        cbytes = client_file.getvalue()
        cwb = load_workbook(BytesIO(cbytes), data_only=False)
        c_sheet = st.selectbox("Client BOQ worksheet", cwb.sheetnames, key="rate_client_sheet")
        cws = cwb[c_sheet]
        c_header = int(st.number_input("Client BOQ header row", min_value=1, max_value=max(1, cws.max_row), value=2, key="rate_client_header"))
        cmap = get_header_map(cws, c_header)
        if not cmap:
            st.error("No column headings found in the selected header row.")
            st.stop()
        clabels = list(cmap)
        cdesc_label = st.selectbox("Client item description column", clabels, index=next((i for i,x in enumerate(clabels) if "description" in x.lower()), 0), key="rate_cdesc")
        cspec_label = st.selectbox("Client specification column (optional)", ["None"] + clabels, index=next((i+1 for i,x in enumerate(clabels) if "specification" in x.lower()), 0), key="rate_cspec")
        cqty_label = st.selectbox("Client quantity column (optional)", ["None"] + clabels, index=next((i+1 for i,x in enumerate(clabels) if "qty" in x.lower() or "quantity" in x.lower()), 0), key="rate_cqty")
        cunit_label = st.selectbox("Client unit column (optional)", ["None"] + clabels, index=next((i+1 for i,x in enumerate(clabels) if "unit" in x.lower()), 0), key="rate_cunit")
        ccode_label = st.selectbox("Client part code column (optional)", ["None"] + clabels, key="rate_ccode")
        supply_label = st.selectbox("Client Supply Rate column", ["None"] + clabels, index=next((i+1 for i,x in enumerate(clabels) if "supply rate" in x.lower()), 0), key="rate_supplycol")
        install_label = st.selectbox("Client Installation Rate column", ["None"] + clabels, index=next((i+1 for i,x in enumerate(clabels) if "installation rate" in x.lower()), 0), key="rate_installcol")
        make_label = st.selectbox("Client Make column (optional)", ["None"] + clabels, index=next((i+1 for i,x in enumerate(clabels) if "make" in x.lower()), 0), key="rate_makecol")

        oem_bytes = oem_file.getvalue()
        ext = os.path.splitext(oem_file.name.lower())[1]
        oem_df = None
        if ext in [".xlsx", ".xls"]:
            try:
                sheets = parse_excel_oem(oem_bytes, oem_file.name)
                if sheets:
                    sheet_names = [x[0] for x in sheets]
                    chosen = st.selectbox("OEM worksheet", sheet_names, key="oem_sheet")
                    _, detected_header, raw_df = next(x for x in sheets if x[0] == chosen)
                    oem_cols = list(raw_df.columns)
                    odesc = st.selectbox("OEM description column", oem_cols, index=next((i for i,x in enumerate(oem_cols) if "description" in x.lower() or "item" in x.lower()), 0), key="oem_desc")
                    ocode = st.selectbox("OEM part code column (optional)", ["None"] + oem_cols, key="oem_code")
                    osupply = st.selectbox("OEM supply rate column", ["None"] + oem_cols, index=next((i+1 for i,x in enumerate(oem_cols) if "supply rate" in x.lower() or ("rate" in x.lower() and "install" not in x.lower())), 0), key="oem_supply")
                    oinstall = st.selectbox("OEM installation rate column", ["None"] + oem_cols, index=next((i+1 for i,x in enumerate(oem_cols) if "installation rate" in x.lower() or "install rate" in x.lower()), 0), key="oem_install")
                    omake = st.selectbox("OEM make / brand column (optional)", ["None"] + oem_cols, key="oem_make")
                    extracted = []
                    for _, row in raw_df.iterrows():
                        desc = clean_text(row.get(odesc, ""))
                        if not desc:
                            continue
                        sr = row.get(osupply) if osupply != "None" else None
                        ir = row.get(oinstall) if oinstall != "None" else None
                        # Ignore section headings and subtotal rows that have no unit rates.
                        if pd.isna(sr) and pd.isna(ir):
                            continue
                        extracted.append({"Description": desc, "Part Code": clean_text(row.get(ocode, "")) if ocode != "None" else "", "Supply Rate": sr, "Installation Rate": ir, "Make": clean_text(row.get(omake, "")) if omake != "None" else ""})
                    oem_df = pd.DataFrame(extracted)
            except Exception as e:
                st.error(f"Could not read OEM Excel: {e}")
        else:
            try:
                extracted_text = ocr_file(oem_bytes, oem_file.name)
                with st.expander("Review OCR text extracted from OEM file", expanded=False):
                    st.text_area("Extracted text", extracted_text, height=220, key="ocr_preview")
                st.warning("OCR rate extraction is an estimate. Check every extracted description and rate below before matching.")
                oem_df = ocr_lines_to_df(extracted_text)
            except Exception as e:
                st.error(str(e))

        if oem_df is not None and not oem_df.empty:
            st.markdown("**Review / correct OEM extracted items and rates**")
            oem_df = st.data_editor(oem_df, num_rows="dynamic", use_container_width=True, key="oem_review_editor")
            if st.button("Compare OEM with Client BOQ", type="primary", key="compare_boq"):
                client_items = find_client_items(cws, c_header, cmap[cdesc_label], cmap[cspec_label] if cspec_label != "None" else None, cmap[cqty_label] if cqty_label != "None" else None, cmap[cunit_label] if cunit_label != "None" else None, cmap[ccode_label] if ccode_label != "None" else None)
                if not client_items:
                    st.error("No client BOQ line items found. Check the selected header and description/quantity columns.")
                else:
                    comparison, extras = make_comparison(client_items, oem_df)
                    st.session_state["rate_comparison"] = comparison
                    st.session_state["rate_extras"] = extras
                    st.session_state["rate_client_bytes"] = cbytes
                    st.session_state["rate_client_sheet"] = c_sheet
                    st.session_state["rate_client_columns"] = {"supply": cmap[supply_label] if supply_label != "None" else None, "install": cmap[install_label] if install_label != "None" else None, "make": cmap[make_label] if make_label != "None" else None}
                    st.session_state["rate_oem_data"] = oem_df
                    st.success(f"Compared {len(client_items)} client items with {len(oem_df)} OEM items.")

        if "rate_comparison" in st.session_state:
            st.markdown("### Client item matching results")
            comp = st.data_editor(st.session_state["rate_comparison"], use_container_width=True, hide_index=True, disabled=["Client Row", "Client Description", "Client Qty", "Client Unit", "OEM Row", "OEM Description", "Match Score", "Status"], key="rate_match_editor", column_config={"Approved": st.column_config.CheckboxColumn("Approve rate fill"), "Supply Rate": st.column_config.NumberColumn("Supply Rate", format="%.2f"), "Installation Rate": st.column_config.NumberColumn("Installation Rate", format="%.2f")})
            st.session_state["rate_comparison"] = comp
            extras = st.session_state.get("rate_extras", pd.DataFrame())
            st.markdown("### Items found only in OEM BOQ (Extra)")
            if extras.empty:
                st.success("No extra OEM items detected.")
            else:
                st.dataframe(extras, use_container_width=True, hide_index=True)
            missing_count = int((comp["Status"] == "Missing in OEM").sum())
            review_count = int((comp["Status"] == "Review").sum())
            approved_count = int(comp["Approved"].fillna(False).sum())
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Matched / Review", len(comp) - missing_count)
            col2.metric("Missing in OEM", missing_count)
            col3.metric("Extra in OEM", len(extras))
            col4.metric("Approved", approved_count)
            if st.button("Apply approved rates and create final client BOQ", type="primary", key="apply_rates"):
                out_wb = load_workbook(BytesIO(st.session_state["rate_client_bytes"]))
                out_ws = out_wb[st.session_state["rate_client_sheet"]]
                cols = st.session_state["rate_client_columns"]
                if not cols["supply"] and not cols["install"]:
                    st.error("Please select at least one client rate column.")
                else:
                    for _, row in comp.iterrows():
                        if not bool(row.get("Approved", False)):
                            continue
                        rr = int(row["Client Row"])
                        qty = row.get("Client Qty")
                        if cols["supply"] and pd.notna(row.get("Supply Rate")):
                            out_ws.cell(rr, cols["supply"]).value = float(row["Supply Rate"])
                        if cols["install"] and pd.notna(row.get("Installation Rate")):
                            out_ws.cell(rr, cols["install"]).value = float(row["Installation Rate"])
                        if cols["make"] and clean_text(row.get("Make")):
                            out_ws.cell(rr, cols["make"]).value = clean_text(row["Make"])
                    # Recalculate line totals when matching total columns are present.
                    header_values = {clean_text(out_ws.cell(c_header, c).value).lower(): c for c in range(1, out_ws.max_column + 1)}
                    qty_col = next((v for k,v in header_values.items() if "qty" in k or "quantity" in k), None)
                    supply_rate_col = cols["supply"]
                    install_rate_col = cols["install"]
                    supply_total_col = next((v for k,v in header_values.items() if "total supply" in k or "supply amount" in k), None)
                    install_total_col = next((v for k,v in header_values.items() if "total installation" in k or "installation amount" in k), None)
                    if qty_col and supply_rate_col and supply_total_col:
                        for r in range(c_header + 1, out_ws.max_row + 1):
                            if out_ws.cell(r, supply_rate_col).value is not None and out_ws.cell(r, qty_col).value is not None:
                                out_ws.cell(r, supply_total_col).value = f"={get_column_letter(supply_rate_col)}{r}*{get_column_letter(qty_col)}{r}"
                    if qty_col and install_rate_col and install_total_col:
                        for r in range(c_header + 1, out_ws.max_row + 1):
                            if out_ws.cell(r, install_rate_col).value is not None and out_ws.cell(r, qty_col).value is not None:
                                out_ws.cell(r, install_total_col).value = f"={get_column_letter(install_rate_col)}{r}*{get_column_letter(qty_col)}{r}"
                    # Add an auditable comparison sheet, preserving original client sheets.
                    if "BOQ Comparison" in out_wb.sheetnames:
                        del out_wb["BOQ Comparison"]
                    cws2 = out_wb.create_sheet("BOQ Comparison")
                    cws2.append(["Client Row", "Client Description", "Client Qty", "Client Unit", "OEM Row", "OEM Description", "Match Score %", "Status", "Supply Rate", "Installation Rate", "Make", "Approved"])
                    for _, row in comp.iterrows():
                        cws2.append([row.get("Client Row"), row.get("Client Description"), row.get("Client Qty"), row.get("Client Unit"), row.get("OEM Row"), row.get("OEM Description"), row.get("Match Score"), row.get("Status"), row.get("Supply Rate"), row.get("Installation Rate"), row.get("Make"), bool(row.get("Approved", False))])
                    cws2.append([])
                    cws2.append(["OEM-only extra items"])
                    cws2.append(["OEM Row", "OEM Description", "Supply Rate", "Installation Rate", "Make", "Status"])
                    for _, row in extras.iterrows():
                        cws2.append([row.get("OEM Row"), row.get("OEM Description"), row.get("Supply Rate"), row.get("Installation Rate"), row.get("Make"), row.get("Status")])
                    style_comparison_sheet(cws2)
                    for col in cws2.columns:
                        letter = get_column_letter(col[0].column)
                        cws2.column_dimensions[letter].width = min(55, max(14, max(len(clean_text(cell.value)) for cell in col if cell.value is not None) + 2))
                    out_wb.calculation.fullCalcOnLoad = True
                    out_wb.calculation.forceFullCalc = True
                    result = BytesIO(); out_wb.save(result)
                    st.success("Final client BOQ created. Please review the BOQ Comparison sheet before issuing it.")
                    st.download_button("Download Final Client BOQ with Comparison", result.getvalue(), "Final_Client_BOQ_with_Comparison.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
        elif oem_df is not None:
            st.warning("No readable OEM item rows were extracted. Try a clearer file/photo or another worksheet.")

st.divider()
st.caption("MEP BOQ Automation | Review all extracted data and approve matches before issuing rates.")