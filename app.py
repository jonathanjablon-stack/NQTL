import streamlit as st
import pandas as pd
from docx import Document
import io
import re
import json

st.set_page_config(page_title="NQTL Assembly Pro", layout="wide")

# [span_6](start_span)[span_7](start_span)[span_8](start_span)[span_9](start_span)[span_10](start_span)Expanded list to catch variations in naming[span_6](end_span)[span_7](end_span)[span_8](end_span)[span_9](end_span)[span_10](end_span)
METRIC_MAP = {
    "totalclaims": "Total Claims Incurred During the Plan Year",
    "deniedbasedonlack": "Denied Based on Lack of Medical Necessity",
    "lackofmedicalnecessityoverturned": "Lack of Medical Necessity Overturned on Appeal",
    "submittedforpriorauth": "Submitted for Prior Authorization",
    "priorauthclaimsdenied": "Prior Authorization Claims Denied Due to Non-Administrative",
    "priorauthoverturned": "Prior Authorization Claims Denied Due to Non-Administrative Reasons Overturned",
    "timeforpriorauthreq": "Time (in Days) for Prior Authorization Requests",
    "timeforpriorauthapp": "Time (in Days) for Prior Authorization Appeals",
    "submittedforconcurrent": "Submitted for Concurrent Review",
    "concurrentdenied": "Concurrent Review Claims Denied Due to Non-Administrative",
    "concurrentoverturned": "Concurrent Review Claims Denied Due to Non-Administrative Reasons Overturned",
    "timeforconcurrentreq": "Time (in Days) for Concurrent Review Requests",
    "timeforconcurrentapp": "Time (in Days) for Concurrent Review Appeals",
    "submittedforretro": "Submitted for Retrospective Review",
    "retrodenied": "Retrospective Review Claims Denied Due to Non-Administrative",
    "retrooverturned": "Retrospective Review Claims Denied Due to Non-Administrative Reasons Overturned",
    "timeforretroreq": "Time (in Days) for Retrospective Review Requests",
    "timeforretroapp": "Time (in Days) for Retrospective Review Appeals"
}

def clean(text):
    return re.sub(r'[^a-zA-Z0-9]', '', str(text)).lower()

def extract_excel(excel_file):
    data = {}
    try:
        xls = pd.ExcelFile(excel_file.getvalue(), engine='calamine')
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet, header=None).astype(str)
            for r_idx, row in df.iterrows():
                row_str = clean(" ".join(row.values))
                for key, full_name in METRIC_MAP.items():
                    if key in row_str:
                        if full_name not in data: data[full_name] = {}
                        # [span_11](start_span)[span_12](start_span)[span_13](start_span)[span_14](start_span)Search window below header[span_11](end_span)[span_12](end_span)[span_13](end_span)[span_14](end_span)
                        for i in range(1, 12):
                            if r_idx + i < len(df):
                                for c_idx, cell in enumerate(df.iloc[r_idx+i]):
                                    c_norm = clean(cell)
                                    # [span_15](start_span)[span_16](start_span)Handle variations like "In-Network" vs "IN"[span_15](end_span)[span_16](end_span)
                                    label = None
                                    if "inpatientin" in c_norm: label = "Inpatient IN"
                                    elif "inpatientoon" in c_norm: label = "Inpatient OON"
                                    elif "outpatientin" in c_norm: label = "Outpatient IN"
                                    elif "outpatientoon" in c_norm: label = "Outpatient OON"
                                    
                                    if label:
                                        # [span_17](start_span)[span_18](start_span)[span_19](start_span)[span_20](start_span)Grab everything to the right[span_17](end_span)[span_18](end_span)[span_19](end_span)[span_20](end_span)
                                        vals = [x if x != "nan" else "" for x in df.iloc[r_idx+i, c_idx+1:c_idx+4].values]
                                        data[full_name][label] = vals
                                        break
    except Exception as e: st.error(f"Excel Error: {e}")
    return data

def inject_word(word_file, data):
    doc = Document(word_file)
    count = 0
    active_metric = None
    
    for table in doc.tables:
        for row in table.rows:
            # [span_21](start_span)[span_22](start_span)[span_23](start_span)Detect Metric Header anywhere in the row[span_21](end_span)[span_22](end_span)[span_23](end_span)
            row_text_clean = clean(" ".join([c.text for c in row.cells]))
            for key, full_name in METRIC_MAP.items():
                if key in row_text_clean:
                    active_metric = full_name
                    break
            
            if not active_metric: continue
            
            # [span_24](start_span)[span_25](start_span)[span_26](start_span)[span_27](start_span)Find the label (Inpatient IN, etc) in any cell[span_24](end_span)[span_25](end_span)[span_26](end_span)[span_27](end_span)
            for idx, cell in enumerate(row.cells):
                c_norm = clean(cell.text)
                target = None
                if "inpatientin" in c_norm: target = "Inpatient IN"
                elif "inpatientoon" in c_norm: target = "Inpatient OON"
                elif "outpatientin" in c_norm: target = "Outpatient IN"
                elif "outpatientoon" in c_norm: target = "Outpatient OON"
                
                if target and target in data[active_metric]:
                    vals = data[active_metric][target]
                    # [span_28](start_span)[span_29](start_span)[span_30](start_span)[span_31](start_span)Fill the 3 cells following the label[span_28](end_span)[span_29](end_span)[span_30](end_span)[span_31](end_span)
                    for i in range(min(len(vals), len(row.cells) - idx - 1)):
                        if vals[i]:
                            row.cells[idx + 1 + i].text = str(vals[i])
                    count += 1
                    break
    return doc, count

st.title("🚀 NQTL Multi-Agent Assembly Engine")
st.info("This version uses fuzzy-logic matching to bypass template changes and merged cells.")

e_file = st.file_uploader("Excel Source", type="xlsx")
w_file = st.file_uploader("Word Template", type="docx")

if e_file and w_file:
    if st.button("Generate Final Analysis", type="primary"):
        extracted = extract_excel(e_file)
        final_doc, updates = inject_word(w_file, extracted)
        
        if updates > 0:
            st.success(f"Successfully processed {updates} data points.")
            buf = io.BytesIO()
            final_doc.save(buf)
            st.download_button("Download Completed .docx", buf.getvalue(), "Final_Analysis.docx")
        else:
            st.error("No data could be mapped. Check Diagnostic Report.")
        
        with st.expander("Diagnostic Report (Copy for troubleshooting)"):
            st.code(json.dumps({"updates": updates, "data": extracted}, indent=2))
