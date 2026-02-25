import streamlit as st
import pandas as pd
from docx import Document
import io

st.set_page_config(page_title="NQTL Document Assembly Engine", layout="centered")

def extract_concurrent_review_data(excel_file):
    """
    Intelligently hunts for the Concurrent Review data in the Excel file,
    regardless of where the client moved the rows or columns.
    """
    try:
        xls = pd.ExcelFile(excel_file)
        # Find the sheet name that contains "Concurrent Review"
        sheet_name = next((s for s in xls.sheet_names if "Concurrent" in s), None)
        
        if not sheet_name:
            st.warning("Could not find a sheet named 'Concurrent Review' in the Excel file.")
            return {}

        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        extracted_data = {}
        
        # Define the metrics we are hunting for
        target_metrics = [
            "Number (#) of Claims Submitted for Concurrent Review",
            "Percentage (%) of Concurrent Review Claims Denied Due to Non-Administrative Reasons",
            "Percentage (%) of Concurrent Review Claims Denied Due to Non-Administrative Reasons Overturned on Appeal",
            "Average Processing Time (in Days) for Concurrent Review Requests",
            "Average Processing Time (in Days) for Concurrent Review Appeals"
        ]

        # Scan the dataframe to find the metrics and extract the 4 rows below them
        for metric in target_metrics:
            extracted_data[metric] = {}
            for row_idx in range(len(df)):
                for col_idx in range(len(df.columns)):
                    cell_value = str(df.iloc[row_idx, col_idx]).strip()
                    if metric.lower() in cell_value.lower():
                        # We found the metric header! Now grab the data for the next 4 rows
                        # Assuming structure: Col = Label (Inpatient IN), Col+1 = M/S, Col+2 = MH, Col+3 = SUD
                        for i in range(1, 5): 
                            if row_idx + i < len(df):
                                row_label = str(df.iloc[row_idx + i, col_idx]).strip()
                                if row_label in ["Inpatient IN", "Inpatient OON", "Outpatient IN", "Outpatient OON"]:
                                    ms_val = str(df.iloc[row_idx + i, col_idx + 1])
                                    mh_val = str(df.iloc[row_idx + i, col_idx + 2])
                                    sud_val = str(df.iloc[row_idx + i, col_idx + 3])
                                    
                                    # Clean up empty cells (nan)
                                    extracted_data[metric][row_label] = {
                                        "M/S": "" if ms_val.lower() == "nan" else ms_val,
                                        "MH": "" if mh_val.lower() == "nan" else mh_val,
                                        "SUD": "" if sud_val.lower() == "nan" else sud_val
                                    }
                        break # Move to next metric once found
        return extracted_data
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return {}

def inject_data_into_word(word_file, client_data):
    """
    Hunts through the Word document for specific text and injects the data.
    """
    doc = Document(word_file)
    tables_updated = 0
    current_metric = None

    for table in doc.tables:
        for row in table.rows:
            try:
                header_text = row.cells[0].text.strip().replace("\n", " ")
                
                # 1. Identify if this row is a Metric Header in our extracted data
                if header_text in client_data and client_data[header_text]:
                    current_metric = header_text
                    continue 
                
                # 2. Identify if this is a Data Row (e.g., "Inpatient IN") under a known Metric
                if current_metric and header_text in client_data[current_metric]:
                    data_for_row = client_data[current_metric][header_text]
                    
                    if len(row.cells) >= 4:
                        # Inject the data directly into the Word table cells
                        row.cells[1].text = data_for_row.get("M/S", "No Data")
                        row.cells[2].text = data_for_row.get("MH", "No Data")
                        row.cells[3].text = data_for_row.get("SUD", "No Data")
                        tables_updated += 1
                
                # Reset metric if we hit a blank row to avoid false positives
                if not header_text:
                    current_metric = None
            except Exception:
                continue

    return doc, tables_updated

# --- FRONT-END UI ---
st.title("📄 NQTL Document Assembly Engine")
st.markdown("Upload a completed client Excel form and your blank Word template. The engine will intelligently map the data and generate a final deliverable.")

st.markdown("### 1. Upload Files")
col1, col2 = st.columns(2)
with col1:
    excel_upload = st.file_uploader("Upload Client Excel Form (.xlsx)", type=["xlsx"])
with col2:
    word_upload = st.file_uploader("Upload Blank Word Template (.docx)", type=["docx"])

if excel_upload and word_upload:
    st.markdown("### 2. Generate Deliverable")
    if st.button("Run Assembly Engine", type="primary"):
        with st.spinner("Analyzing Excel data and mapping to Word document..."):
            
            # Step 1: Extract Data
            extracted_data = extract_concurrent_review_data(excel_upload)
            
            # Step 2: Inject Data
            if extracted_data:
                final_doc, updates = inject_data_into_word(word_upload, extracted_data)
                
                # Step 3: Prepare for Download
                output_stream = io.BytesIO()
                final_doc.save(output_stream)
                output_stream.seek(0)
                
                st.success(f"✅ Success! Injected data into {updates} rows in the Word document.")
                
                st.download_button(
                    label="⬇️ Download Completed NQTL Analysis",
                    data=output_stream,
                    file_name="Completed_NQTL_Analysis.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
