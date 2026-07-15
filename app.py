import streamlit as st
import json
import pandas as pd
from pypdf import PdfReader
import io

# We only map keys that lead directly to your 10 required fields.
# Anything else in the PDF will be automatically ignored and dropped.
FIELD_MAP = {
    # --- FIRST NAME ---
    "First name": "First name",
    "Personal_name_first~7": "First name",
    "Personal_name_first~8": "First name",
    "Personal_name_first~9": "First name",
    "EF_Emp_name_first~4": "First name",
    
    # --- LAST NAME ---
    "Last name": "Last name",
    "Personal_name_last~7": "Last name",
    "Personal_name_last~8": "Last name",
    "Personal_name_last~9": "Last name",
    "EF_Emp_name_last~4": "Last name",
    
    # --- PHONE ---
    "Phone": "Phone",
    "EF_Emp_phone_primary~3": "Phone",
    "EF_Emp_phone_primary~4": "Phone",
    "EF_Emp_phone_primary~7": "Phone",
    
    # --- DOB ---
    "DOB": "DOB",
    "EF_Emp_birth_date~4": "DOB",
    "EF_Emp_birth_date~7": "DOB",
    
    # --- ADDRESS STREET ---
    "Address Street": "Address Street",
    "ResAddr_addr__street_1~8": "Address Street",
    "ResAddr_addr__street_1~9": "Address Street",
    "ResAddrTran_addr__street_full_VC~7": "Address Street",
    "EF_Emp_Residence_street_1~3": "Address Street",
    "EF_Emp_Residence_street_1~4": "Address Street",
    
    # --- ADDRESS CITY ---
    "Address City": "Address City",
    "ResAddr_addr__city~8": "Address City",
    "ResAddr_addr__city~9": "Address City",
    "EF_Emp_Residence_city~3": "Address City",
    "EF_Emp_Residence_city~4": "Address City",
    
    # --- ADDRESS STATE ---
    "Address State": "Address State",
    "ResAddr_addr__state_desc~8": "Address State",
    "ResAddr_addr__state_desc~9": "Address State",
    "EF_Emp_Residence_state_rdo~3": "Address State",
    "EF_Emp_Residence_state_rdo~4": "Address State",
    
    # --- ADDRESS ZIP ---
    "Address Zip": "Address Zip",
    "ResAddr_addr__zip_code~8": "Address Zip",
    "ResAddr_addr__zip_code~9": "Address Zip",
    "EF_Emp_Residence_zip_code~3": "Address Zip",
    "EF_Emp_Residence_zip_code~4": "Address Zip",
    
    # --- SSN ---
    "SSN": "SSN",
    "Personal_ssn~7": "SSN",
    "Personal_ssn~8": "SSN",
    "Personal_ssn~9": "SSN",
    "EF_Emp_ssn~4": "SSN",
    
    # --- EMAIL ---
    "Email Address": "Email Address",
    "EF_Emp_email~4": "Email Address",
    "EF_Emp_email~7": "Email Address",
}

# The ONLY columns that will exist in the final Excel file
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
                    
                    # Convert raw PDF key to our clean Excel target name
                    clean_column_name = FIELD_MAP.get(field_name, field_name)
                    
                    # Only grab values that map to our final list
                    if clean_column_name in FINAL_COLUMNS:
                        if value:
                            # Keep longest value to prevent page-overwrite issues
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
            # === AUTO-SAVE JUST THESE 10 FIELDS TO GOOGLE COLAB FILE SYSTEM ===
            # Creates a clean extracted JSON containing only our 10 targeted fields
            clean_json_data = {}
            for filename, data in all_form_data.items():
                clean_json_data[filename] = {col: data.get(col, "") for col in FINAL_COLUMNS}
                
            try:
                with open("extracted_data.json", "w", encoding="utf-8") as f:
                    json.dump(clean_json_data, f, indent=4)
            except Exception:
                pass

            # Create DataFrame
            df = pd.DataFrame.from_dict(all_form_data, orient="index")
            
            # This strictly keeps ONLY the 10 fields you want and ignores everything else
            df = df.reindex(columns=FINAL_COLUMNS)
            df.index.name = "Source File Name"
            df = df.fillna("")
            
            st.success(f"🎉 Successfully processed {len(all_form_data)} PDFs!")
            
            # --- VISUAL PREVIEW SECTION ---
            st.subheader("📊 Spreadsheet Preview")
            st.data_editor(df, use_container_width=True, key="excel_preview", disabled=True)
            
            # Save Excel workbook file to memory buffer
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=True, sheet_name='Extracted Data')
            
            st.markdown("---")
            
            # UI Downloads Columns
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📥 Download Excel File",
                    data=buffer.getvalue(),
                    file_name="background_checks_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            with col2:
                st.download_button(
                    label="📥 Download Raw JSON File",
                    data=json.dumps(clean_json_data, indent=4),
                    file_name="background_checks_output.json",
                    mime="application/json"
                )
        else:
            st.warning("No core matching fields were found in the uploaded PDFs.")
