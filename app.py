import streamlit as st
import pandas as pd
from docx import Document
import io
import re
import json

# --- CONFIGURATION & UI SETUP ---
st.set_page_config(page_title="NQTL Document Assembly", layout="wide", page_icon="📄")

TARGET_METRICS = [
    "Total Claims Incurred During the Plan Year",
    "Denied Based on Lack of Medical Necessity",
    "Lack of Medical Necessity Overturned on Appeal",
    "Submitted for Prior Authorization",
    "Prior Authorization Claims Denied Due to Non-Administrative",
    "Prior Authorization Claims Denied Due to Non-Administrative Reasons Overturned",
    "Time (in Days) for Prior Authorization Requests",
    "Time (in Days) for Prior Authorization Appeals",
    "Submitted for Concurrent Review",
    "Concurrent Review Claims Denied Due to Non-Administrative",
    "Concurrent Review Claims Denied Due to Non-Administrative Reasons Overturned",
    "Time (in Days) for Concurrent Review Requests",
    "Time (in Days) for Concurrent Review Appeals",
    "Submitted for Retrospective Review",
    "Retrospective Review Claims Denied Due to Non-Administrative",
    "Retrospective Review Claims Denied Due to Non-Administrative Reasons Overturned",
    "Time (in Days) for Retrospective Review Requests",
    "Time (in Days) for Retrospective Review Appeals"
]

def normalize_text(text):
    """Strips all spaces and special characters for bulletproof matching."""
    if not isinstance(text, str):
        return ""
    return re.sub(r'[^a-zA-Z0-9]', '', text).lower()

def extract_all_data(excel_file):
    """X-Ray Scanner: Finds data regardless of which column the client used."""
    extracted_data = {}
    try:
        xls = pd.ExcelFile(excel_file)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            if df.empty: continue
            
            num_rows, num_cols = df.shape
            for row_idx in range(num_rows):
                row_text = " ".join([str(x) for x in df.iloc[row_idx, :].values])
                
                for metric in TARGET_METRICS:
                    metric_norm = normalize_text(metric)
                    if metric_norm in normalize_text(row_text) and metric_norm != "":
                        if metric not in extracted_data:
                            extracted_data[metric] = {}
                        
                        for i in range(1, 15):
                            if row_idx + i < num_rows:
                                for col_idx in range(num_cols):
                                    cell_val = str(df.iloc[row_idx + i, col_idx]).strip()
                                    norm_val = normalize_text(cell_val)
                                    target_labels = ["inpatientin", "inpatientoon", "outpatientin", "outpatientoon"]
                                    
                                    if norm_val in target_labels:
                                        vals = []
                                        for v_col in range(1, 4):
                                            if col_idx + v_col < num_cols:
                                                val = str(df.iloc[row_idx + i, col_idx + v_col]).strip()
                                                vals.append("" if val.lower() == "nan" else val)
                                            else:
                                                vals.append("")
                                        
                                        clean_label = "Inpatient IN" if norm_val == "inpatientin" else \
                                                      "Inpatient OON" if norm_val == "inpatientoon" else \
                                                      "Outpatient IN" if norm_val == "outpatientin" else "Outpatient OON"
                                                      
                                        extracted_data[metric][clean_label] = vals
                                        break
    except Exception as e:
        st.error(f"Excel Extraction Error: {e}")
    return extracted_data

def inject_data_into_word(word_file, client_data):
    """Injects data into Word, handling side-by-side column layouts."""
    doc = Document(word_file)
    tables_updated = 0
    
    for table in doc.tables:
        current_metric = None
        for row in table.rows:
            try:
                col0_text = row.cells[0].text.strip()
                col0_norm = normalize_text(col0_text)
                
                if col0_norm:
                    matched_metric = next((m for m in client_data.keys() if normalize_text(m) in col0_norm), None)
                    if matched_metric:
                        current_metric = matched_metric
                
                if len(row.cells) >= 4:
                    label_text = row.cells[1].text.strip()
                    if not label_text: 
                        label_text = row.cells[0].text.strip()
                        
                    label_norm = normalize_text(label_text)
                    
                    if current_metric:
                        dict_key = next((k for k in client_data[current_metric].keys() if normalize_text(k) == label_norm), None)
                        
                        if dict_key:
                            data_vals = client_data[current_metric][dict_key]
                            target_col = 2 if len(row.cells) >= 5 else 1
                            
                            cells_injected = False
                            for i, val in enumerate(data_vals):
                                if val and (target_col + i) < len(row.cells):
                                    row.cells[target_col + i].text = val
                                    cells_injected = True
                                    
                            if cells_injected:
                                tables_updated += 1
            except Exception:
                continue

    return doc, tables_updated

# --- FRONT-END UI ---
st.title("📄 NQTL Document Assembly Engine")
st.markdown("Automated mapping of structured Excel inputs to the maintainable Word deliverable.")
st.divider()

col1, col2 = st.columns(2)
with col1:
    st.markdown("#### 1. Upload Excel Data")
    excel_upload = st.file_uploader("Select completed Information Request Form (.xlsx)", type=["xlsx"])
with col2:
    st.markdown("#### 2. Upload Word Template")
    word_upload = st.file_uploader("Select blank NQTL Comparative Analysis (.docx)", type=["docx"])

st.divider()

if excel_upload and word_upload:
    col_center = st.columns([1, 2, 1])[1]
    with col_center:
        run_btn = st.button("🚀 Generate Final Deliverable", use_container_width=True, type="primary")

    if run_btn:
        with st.spinner("Processing documents... please wait."):
            extracted_data = extract_all_data(excel_upload)
            final_doc, updates = inject_data_into_word(word_upload, extracted_data)
            
            # --- RESULTS UI ---
            if updates > 0:
                st.success(f"✅ Success! Successfully mapped {updates} rows of data into the Word document.")
                output_stream = io.BytesIO()
                final_doc.save(output_stream)
                output_stream.seek(0)
                st.download_button(
                    label="⬇️ Download Completed Analysis", 
                    data=output_stream, 
                    file_name="Completed_NQTL.docx", 
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            else:
                st.warning("⚠️ Files processed, but 0 rows were updated in Word.")
            
            # --- DIAGNOSTIC REPORT FOR EASY COPYING ---
            st.markdown("### 🛠️ Developer Diagnostic Report")
            st.info("Hover over the top-right corner of the code block below to copy the results.")
            
            diagnostic_output = {
                "Word_Tables_Updated": updates,
                "Excel_Data_Extracted": extracted_data
            }
            
            # Display as formatted JSON with native copy button
            st.code(json.dumps(diagnostic_output, indent=2), language="json")
 
