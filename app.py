import streamlit as st
import pandas as pd
from docx import Document
import io

st.set_page_config(page_title="NQTL Document Assembly Engine", layout="centered")

# The master list of Word tables the engine will hunt for across all Excel sheets
TARGET_METRICS = [
    # Medical Management
    "Number (#) of Total Claims Incurred During the Plan Year",
    "Percentage (%) of Claims Denied Based on Lack of Medical Necessity",
    "Percentage (%) of Claims Denied Based on Lack of Medical Necessity Overturned on Appeal",
    # Prior Authorization
    "Number (#) of Claims Submitted for Prior Authorization",
    "Percentage (%) of Prior Authorization Claims Denied Due to Non-Administrative Reasons",
    "Percentage (%) of Prior Authorization Claims Denied Due to Non-Administrative Reasons Overturned on Appeal",
    "Average Processing Time (in Days) for Prior Authorization Requests",
    "Average Processing Time (in Days) for Prior Authorization Appeals",
    # Concurrent Review
    "Number (#) of Claims Submitted for Concurrent Review",
    "Percentage (%) of Concurrent Review Claims Denied Due to Non-Administrative Reasons",
    "Percentage (%) of Concurrent Review Claims Denied Due to Non-Administrative Reasons Overturned on Appeal",
    "Average Processing Time (in Days) for Concurrent Review Requests",
    "Average Processing Time (in Days) for Concurrent Review Appeals",
    # Retrospective Review
    "Number (#) of Claims Submitted for Retrospective Review",
    "Percentage (%) of Retrospective Review Claims Denied Due to Non-Administrative Reasons",
    "Percentage (%) of Retrospective Review Claims Denied Due to Non-Administrative Reasons Overturned on Appeal",
    "Average Processing Time (in Days) for Retrospective Review Requests",
    "Average Processing Time (in Days) for Retrospective Review Appeals",
    # Experimental/Investigational
    "Percentage (%) of Claims Denied as Experimental/Investigational",
    "Average Processing Time (in Days) for Experimental/Investigational Requests",
    "Percentage (%) of Experimental/Investigational Claims Appealed",
    "Percentage (%) of Experimental/Investigational Denials Overturned on Appeal",
    "Average Processing Time (in Days) for Experimental/Investigational Appeals"
]

def extract_all_data(excel_file):
    """
    Crash-proof scanner that reads every sheet in the Excel file.
    """
    extracted_data = {}
    try:
        xls = pd.ExcelFile(excel_file)
        
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            
            # Skip completely empty sheets
            if df.empty:
                continue
                
            num_rows, num_cols = df.shape
            
            # Safely scan the first two columns (if they exist)
            cols_to_scan = min(2, num_cols)
            
            for row_idx in range(num_rows):
                for col_idx in range(cols_to_scan): 
                    cell_value = str(df.iloc[row_idx, col_idx]).strip()
                    
                    for metric in TARGET_METRICS:
                        if metric.lower() in cell_value.lower():
                            if metric not in extracted_data:
                                extracted_data[metric] = {}
                            
                            # We found a table header! Scan the rows below it for data
                            for i in range(1, 15):
                                if row_idx + i < num_rows:
                                    row_label = str(df.iloc[row_idx + i, col_idx]).strip()
                                    
                                    # Skip empty rows or standard sub-headers
                                    ignore_list = ["nan", "m/s", "mh", "sud", "medical/surgical", "mental health", "substance use disorder"]
                                    if not row_label or row_label.lower() in ignore_list:
                                        continue
                                        
                                    # Grab the next 4 columns of data safely
                                    vals = []
                                    for v_col in range(1, 5):
                                        if col_idx + v_col < num_cols:
                                            val = str(df.iloc[row_idx + i, col_idx + v_col]).strip()
                                            vals.append("" if val.lower() == "nan" else val)
                                        else:
                                            vals.append("")
                                            
                                    extracted_data[metric][row_label] = vals
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        
    return extracted_data

def inject_data_into_word(word_file, client_data):
    """
    Hunts through the Word document for matching tables and injects the data.
    """
    doc = Document(word_file)
    tables_updated = 0
    
    for table in doc.tables:
        current_metric = None
        
        for row in table.rows:
            try:
                header_text = row.cells[0].text.strip().replace("\n", " ")
                
                # Check if this Word row matches a Metric Header we extracted
                matched_metric = next((m for m in client_data.keys() if m.lower() in header_text.lower()), None)
                
                if matched_metric:
                    current_metric = matched_metric
                    continue
                    
                # Check if this is a Data Row under a known metric
                if current_metric and header_text in client_data[current_metric]:
                    data_vals = client_data[current_metric][header_text]
                    
                    cells_injected = False
                    for i in range(min(len(data_vals), len(row.cells) - 1)):
                        if data_vals[i]: 
                            row.cells[i+1].text = data_vals[i]
                            cells_injected = True
                            
                    if cells_injected:
                        tables_updated += 1
                        
                if not header_text:
                    current_metric = None
                    
            except Exception:
                continue

    return doc, tables_updated

# --- FRONT-END UI ---
st.title("📄 NQTL Document Assembly Engine")
st.markdown("Upload a completed client Excel form and your blank Word template.")

col1, col2 = st.columns(2)
with col1:
    excel_upload = st.file_uploader("Upload Client Excel Form (.xlsx)", type=["xlsx"])
with col2:
    word_upload = st.file_uploader("Upload Blank Word Template (.docx)", type=["docx"])

if excel_upload and word_upload:
    if st.button("Run Assembly Engine", type="primary"):
        with st.spinner("Scanning Excel sheets and mapping to Word document..."):
            
            extracted_data = extract_all_data(excel_upload)
            
            if extracted_data:
                final_doc, updates = inject_data_into_word(word_upload, extracted_data)
                
                output_stream = io.BytesIO()
                final_doc.save(output_stream)
                output_stream.seek(0)
                
                st.success(f"✅ Success! Injected data into {updates} rows across the document.")
                
                st.download_button(
                    label="⬇️ Download Completed NQTL Analysis",
                    data=output_stream,
                    file_name="Completed_NQTL_Analysis.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.warning("Could not find any matching NQTL data tables in the uploaded Excel file. Make sure the text headers in the Excel file match the target list.")
