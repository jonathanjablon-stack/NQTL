import streamlit as st
import pandas as pd
from docx import Document
import io

st.set_page_config(page_title="NQTL Document Assembly Engine", layout="centered")

TARGET_METRICS = [
    "Number (#) of Total Claims Incurred During the Plan Year",
    "Percentage (%) of Claims Denied Based on Lack of Medical Necessity",
    "Percentage (%) of Claims Denied Based on Lack of Medical Necessity Overturned on Appeal",
    "Number (#) of Claims Submitted for Prior Authorization",
    "Percentage (%) of Prior Authorization Claims Denied Due to Non-Administrative Reasons",
    "Percentage (%) of Prior Authorization Claims Denied Due to Non-Administrative Reasons Overturned on Appeal",
    "Average Processing Time (in Days) for Prior Authorization Requests",
    "Average Processing Time (in Days) for Prior Authorization Appeals",
    "Number (#) of Claims Submitted for Concurrent Review",
    "Percentage (%) of Concurrent Review Claims Denied Due to Non-Administrative Reasons",
    "Percentage (%) of Concurrent Review Claims Denied Due to Non-Administrative Reasons Overturned on Appeal",
    "Average Processing Time (in Days) for Concurrent Review Requests",
    "Average Processing Time (in Days) for Concurrent Review Appeals",
    "Number (#) of Claims Submitted for Retrospective Review",
    "Percentage (%) of Retrospective Review Claims Denied Due to Non-Administrative Reasons",
    "Percentage (%) of Retrospective Review Claims Denied Due to Non-Administrative Reasons Overturned on Appeal",
    "Average Processing Time (in Days) for Retrospective Review Requests",
    "Average Processing Time (in Days) for Retrospective Review Appeals",
    "Percentage (%) of Claims Denied as Experimental/Investigational",
    "Average Processing Time (in Days) for Experimental/Investigational Requests",
    "Percentage (%) of Experimental/Investigational Claims Appealed",
    "Percentage (%) of Experimental/Investigational Denials Overturned on Appeal",
    "Average Processing Time (in Days) for Experimental/Investigational Appeals"
]

def extract_all_data(excel_file):
    """
    Scans Excel sheets looking for side-by-side or stacked data configurations.
    """
    extracted_data = {}
    try:
        xls = pd.ExcelFile(excel_file)
        
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            if df.empty:
                continue
                
            num_rows, num_cols = df.shape
            
            for row_idx in range(num_rows):
                col0_val = str(df.iloc[row_idx, 0]).strip()
                
                for metric in TARGET_METRICS:
                    if metric.lower() in col0_val.lower():
                        if metric not in extracted_data:
                            extracted_data[metric] = {}
                        
                        # Scan the current row and the next 8 rows for our data labels
                        for i in range(0, 10):
                            if row_idx + i < num_rows:
                                label_col0 = str(df.iloc[row_idx + i, 0]).strip()
                                label_col1 = str(df.iloc[row_idx + i, 1]).strip() if num_cols > 1 else ""
                                
                                # The label ("Inpatient IN") might be in col 0 or col 1
                                target_labels = ["Inpatient IN", "Inpatient OON", "Outpatient IN", "Outpatient OON"]
                                row_label = label_col1 if label_col1 in target_labels else label_col0
                                
                                if row_label in target_labels:
                                    val_start_col = 1 if row_label == label_col0 else 2
                                    
                                    vals = []
                                    for v_col in range(3): # Grab M/S, MH, SUD
                                        c_idx = val_start_col + v_col
                                        if c_idx < num_cols:
                                            val = str(df.iloc[row_idx + i, c_idx]).strip()
                                            vals.append("" if val.lower() == "nan" else val)
                                        else:
                                            vals.append("")
                                    
                                    extracted_data[metric][row_label] = vals
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        
    return extracted_data

def inject_data_into_word(word_file, client_data):
    """
    Injects data into Word tables accounting for the side-by-side layout.
    """
    doc = Document(word_file)
    tables_updated = 0
    
    for table in doc.tables:
        current_metric = None
        
        for row in table.rows:
            try:
                # In your Word doc, Column 0 is the Metric, Column 1 is the Label (Inpatient IN)
                col0_text = row.cells[0].text.strip().replace("\n", " ")
                
                # Check if Column 0 contains a target metric
                if col0_text:
                    matched_metric = next((m for m in client_data.keys() if m.lower() in col0_text.lower()), None)
                    if matched_metric:
                        current_metric = matched_metric
                
                # Now check Column 1 for the data label (Inpatient IN, etc.)
                if len(row.cells) >= 5:
                    row_label = row.cells[1].text.strip().replace("\n", " ")
                    
                    if current_metric and row_label in client_data[current_metric]:
                        data_vals = client_data[current_metric][row_label]
                        
                        # Inject into Columns 2 (M/S), 3 (MH), and 4 (SUD)
                        cells_injected = False
                        if data_vals[0]: 
                            row.cells[2].text = data_vals[0]
                            cells_injected = True
                        if len(data_vals) > 1 and data_vals[1]: 
                            row.cells[3].text = data_vals[1]
                        if len(data_vals) > 2 and data_vals[2]: 
                            row.cells[4].text = data_vals[2]
                            
                        if cells_injected:
                            tables_updated += 1
                            
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
                
                if updates > 0:
                    st.success(f"✅ Success! Injected data into {updates} rows across the document.")
                else:
                    st.warning("⚠️ Files were processed, but 0 rows were updated. Check if the Word table structures match the Excel headers.")
                
                st.download_button(
                    label="⬇️ Download Completed NQTL Analysis",
                    data=output_stream,
                    file_name="Completed_NQTL_Analysis.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.error("Could not extract any data from the Excel file.")
