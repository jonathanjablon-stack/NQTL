import streamlit as st
import pandas as pd
from docx import Document
import io
import re

st.set_page_config(page_title="NQTL Document Assembly Engine", layout="wide")

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
    extracted_data = {}
    try:
        xls = pd.ExcelFile(excel_file)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            if df.empty: continue
            
            num_rows, num_cols = df.shape
            for row_idx in range(num_rows):
                for col_idx in range(min(2, num_cols)):
                    cell_value = str(df.iloc[row_idx, col_idx]).strip()
                    cell_norm = normalize_text(cell_value)
                    
                    for metric in TARGET_METRICS:
                        metric_norm = normalize_text(metric)
                        if metric_norm in cell_norm and metric_norm != "":
                            if metric not in extracted_data:
                                extracted_data[metric] = {}
                            
                            for i in range(1, 15):
                                if row_idx + i < num_rows:
                                    label_col0 = str(df.iloc[row_idx + i, 0]).strip()
                                    label_col1 = str(df.iloc[row_idx + i, 1]).strip() if num_cols > 1 else ""
                                    
                                    target_labels = ["inpatientin", "inpatientoon", "outpatientin", "outpatientoon"]
                                    norm_l0 = normalize_text(label_col0)
                                    norm_l1 = normalize_text(label_col1)
                                    
                                    if norm_l0 in target_labels or norm_l1 in target_labels:
                                        val_start_col = 1 if norm_l0 in target_labels else 2
                                        real_label = label_col0 if norm_l0 in target_labels else label_col1
                                        
                                        vals = []
                                        for v_col in range(3):
                                            c_idx = val_start_col + v_col
                                            if c_idx < num_cols:
                                                val = str(df.iloc[row_idx + i, c_idx]).strip()
                                                vals.append("" if val.lower() == "nan" else val)
                                            else:
                                                vals.append("")
                                        
                                        # Only add if there's actual data
                                        if any(vals):
                                            extracted_data[metric][real_label] = vals
    except Exception as e:
        st.error(f"Excel Extraction Error: {e}")
    return extracted_data

def inject_data_into_word(word_file, client_data):
    doc = Document(word_file)
    tables_updated = 0
    
    for table in doc.tables:
        current_metric = None
        for row in table.rows:
            try:
                col0_text = row.cells[0].text.strip()
                col0_norm = normalize_text(col0_text)
                
                # Check if Column 0 contains a target metric
                if col0_norm:
                    matched_metric = next((m for m in client_data.keys() if normalize_text(m) in col0_norm), None)
                    if matched_metric:
                        current_metric = matched_metric
                
                if len(row.cells) >= 4:
                    # Look for the label in col 0 or col 1
                    label_text = row.cells[1].text.strip()
                    if not label_text: 
                        label_text = row.cells[0].text.strip()
                        
                    label_norm = normalize_text(label_text)
                    
                    if current_metric:
                        # Find the matching label in our dictionary using normalized text
                        dict_key = next((k for k in client_data[current_metric].keys() if normalize_text(k) == label_norm), None)
                        
                        if dict_key:
                            data_vals = client_data[current_metric][dict_key]
                            
                            # Assume standard Word layout: Col 0/1 are labels, Col 2,3,4 are data
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
st.title("📄 NQTL Document Assembly Engine (Diagnostic Mode)")

col1, col2 = st.columns(2)
with col1:
    excel_upload = st.file_uploader("Upload Client Excel Form (.xlsx)", type=["xlsx"])
with col2:
    word_upload = st.file_uploader("Upload Blank Word Template (.docx)", type=["docx"])

if excel_upload and word_upload:
    if st.button("Run Assembly Engine", type="primary"):
        with st.spinner("Scanning files..."):
            
            # 1. Test Excel Extraction
            extracted_data = extract_all_data(excel_upload)
            
            st.markdown("### Step 1: Excel Extraction Results")
            if extracted_data:
                st.success("Successfully found data in the Excel file!")
                st.json(extracted_data) # <--- THIS SHOWS YOU EXACTLY WHAT IT FOUND
                
                # 2. Test Word Injection
                final_doc, updates = inject_data_into_word(word_upload, extracted_data)
                
                st.markdown("### Step 2: Word Injection Results")
                if updates > 0:
                    st.success(f"✅ Success! Injected data into {updates} rows across the Word document.")
                    output_stream = io.BytesIO()
                    final_doc.save(output_stream)
                    output_stream.seek(0)
                    st.download_button("⬇️ Download Completed Analysis", data=output_stream, file_name="Completed_NQTL.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                else:
                    st.error("⚠️ The Excel data was read perfectly, but 0 rows were updated in Word. This means the Word table structures do not match the expected format.")
            else:
                st.error("❌ Failed to extract any relevant data from the Excel file. The script could not find the headers.")
