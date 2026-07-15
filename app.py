import streamlit as st
import json
import pandas as pd
from pypdf import PdfReader
import io

# Unified field map that merges keys from both old and new form variants
# directly into your 10 target columns.
FIELD_MAP = {
    # --- FIRST NAME ---
    "First name": "First name",
    "Personal_name_first~7": "First name",
    "Personal_name_first~8": "First name",
    "Personal_name_first~9": "First name",
    "EF_Emp_name_first~4": "First name",
    "EF_Emp_name_first~5": "First name", # New Form variant
    
    # --- LAST NAME ---
    "Last name": "Last name",
    "Personal_name_last~7": "Last name",
    "Personal_name_last~8": "Last name",
    "Personal_name_last~9": "Last name",
    "EF_Emp_name_last~4": "Last name",
    "EF_Emp_name_last~5": "Last name", # New Form variant
    
    # --- PHONE ---
    "Phone": "Phone",
    "EF_Emp_phone_primary~3": "Phone",
    "EF_Emp_phone_primary~4": "Phone",
    "EF_Emp_phone_primary~7": "Phone",
    "EF_Emp_phone_secondary~5": "Phone", # New Form variant
    
    # --- DOB ---
    "DOB": "DOB",
    "EF_Emp_birth_date~4": "DOB",
    "EF_Emp_birth_date~5": "DOB", # New Form variant
    "EF_Emp_birth_date~7": "DOB",
    
    # --- ADDRESS STREET ---
    "Address Street": "Address Street",
    "ResAddr_addr__street_1~8": "Address Street",
    "ResAddr_addr__street_1~9": "Address Street",
    "ResAddrTran_addr__street_full_VC~7": "Address Street",
    "EF_Emp_Residence_street_1~3": "Address Street",
    "EF_Emp_Residence_street_1~4": "Address Street",
    "EF_Emp_Residence_street_1~5": "Address Street", # New Form variant
    
    # --- ADDRESS CITY ---
    "Address City": "Address City",
    "ResAddr_addr__city~8": "Address City",
    "ResAddr_addr__city~9": "Address City",
    "EF_Emp_Residence_city~3": "Address City",
    "EF_Emp_Residence_city~4": "Address City",
    "EF_Emp_Residence_city~5": "Address City", # New Form variant
    
    # --- ADDRESS STATE ---
    "Address State": "Address State",
    "ResAddr_addr__state_desc~8": "Address State",
    "ResAddr_addr__state_desc~9": "Address State",
    "EF_Emp_Residence_state_rdo~3": "Address State",
    "EF_Emp_Residence_state_rdo~4": "Address State",
    "EF_Emp_Residence_state_rdo~5": "Address State", # New Form variant
    
    # --- ADDRESS ZIP ---
    "Address Zip": "Address Zip",
    "ResAddr_addr__zip_code~8": "Address Zip",
    "ResAddr_addr__zip_code~9": "Address Zip",
    "EF_Emp_Residence_zip_code~3": "Address Zip",
    "EF_Emp_Residence_zip_code~4": "Address Zip",
    "EF_Emp_Residence_zip_code~5": "Address Zip", # New Form variant (Safe coverage)
    
    # --- SSN ---
    "SSN": "SSN",
    "Personal_ssn~7": "SSN",
    "Personal_ssn~8": "SSN",
    "Personal_ssn~9": "SSN",
    "EF_Emp_ssn~4": "SSN",
    "EF_Emp_ssn~5": "SSN", # New Form variant
    
    # --- EMAIL ADDRESS ---
    "Email Address": "Email Address",
    "EF_Emp_email~4": "Email Address",
    "EF_Emp_email~5": "Email Address", # New Form variant
    "EF_Emp_email~7": "Email Address",
}

# The strictly filtered columns displayed in Streamlit & written to Excel
FINAL_COLUMNS = [
    "First name", 
    "Last name", 
    "Phone", 
    "DOB", 
    "Address Street", 
    "Address City", 
    "Address State", 
    "Address Zip", 
    "SSN", 
    "Email Address"
]

st.set_page_config(page_title="PDF to Excel Converter", page_icon="📝", layout="centered")

st.title("📝 PDF Background Check Extractor")
st.write("Upload your PDF forms below. You can drag and drop multiple files separately. When you are ready, click **Convert to Excel**.")

with st.form(key="pdf_upload_form", clear_on_submit=False):
    uploaded_files = st.file_uploader(
        "Drop PDF files here", 
        type="pdf", 
        accept_multiple_files=True
    )
    submit_button = st.form_submit_button(label="⚡ Convert to Excel")

if submit_button:
    if not uploaded_files:
        st.error("Please upload at least one PDF file before clicking convert.")
    else:
        all_form_data = {}
        progress_bar = st.progress(0)
        
        for index, uploaded_file in enumerate(uploaded_files):
            try:
                reader = PdfReader(uploaded_file)
                fields = reader.get_fields()
                
                if not fields:
                    continue
                    
                file_data = {}
                for field_name, field_info in fields.items():
                    value = field_info.get("/V", "")
                    if isinstance(value, str):
                        value = value.strip()
                        if value.startswith("/"):
                            value = value.lstrip("/")
                    
                    # Convert raw PDF key to our target spreadsheet name if it exists in map
                    clean_column_name = FIELD_MAP.get(field_name, field_name)
                    
                    # Overwrite protection logic: keeps the longest value
                    if value:
                        if clean_column_name in file_data and file_data[clean_column_name]:
                            if len(str(value)) > len(str(file_data[clean_column_name])):
                                file_data[clean_column_name] = value
                        else:
                            file_data[clean_column_name] = value
                    else:
                        if clean_column_name not in file_data:
                            file_data[clean_column_name] = ""
                    
                all_form_data[uploaded_file.name] = file_data
                
            except Exception as e:
                st.error(f"Error reading {uploaded_file.name}: {e}")
            
            progress_bar.progress((index + 1) / len(uploaded_files))

        if all_form_data:
            # === AUTO-SAVE ALL EXTRACTED RAW FIELDS TO CLOUD DISK ===
            try:
                with open("extracted_data.json", "w", encoding="utf-8") as f:
                    json.dump(all_form_data, f, indent=4)
            except Exception:
                pass

            # Create DataFrame
            df = pd.DataFrame.from_dict(all_form_data, orient="index")
            
            # CRITICAL: Reindex strictly keeps ONLY the 10 columns you asked for in Excel
            df = df.reindex(columns=FINAL_COLUMNS)
            df.index.name = "Source File Name"
            df = df.fillna("")
            
            st.success(f"🎉 Successfully processed {len(all_form_data)} PDFs!")
            
            # --- VISUAL PREVIEW SECTION ---
            st.subheader("📊 Spreadsheet Preview")
            st.write("Review the extracted 10 core fields below:")
            st.data_editor(df, use_container_width=True, key="excel_preview", disabled=True)
            
            # Save Excel workbook file (strictly containing the 10 fields) to memory buffer
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=True, sheet_name='Extracted Data')
            
            st.markdown("---")
            
            # UI Downloads Columns
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📥 Download Excel File (10 Fields Only)",
                    data=buffer.getvalue(),
                    file_name="background_checks_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            with col2:
                # This download delivers the ENTIRE parsed PDF database, raw keys and all!
                st.download_button(
                    label="📥 Download Raw JSON File (Entire PDF Data)",
                    data=json.dumps(all_form_data, indent=4),
                    file_name="background_checks_complete_dump.json",
                    mime="application/json"
                )
        else:
            st.warning("No interactive form fields were found in the uploaded PDFs.")
