%%writefile app.py
import streamlit as st
import json
import pandas as pd
from pypdf import PdfReader
import io

st.set_page_config(page_title="PDF to Excel Converter", page_icon="📄", layout="centered")

st.title("📄 PDF Background Check Extractor")
st.write("Upload your PDF forms below. You can drag and drop multiple files separately. When you are ready, click **Convert to Excel**.")

with st.form(key="pdf_upload_form", clear_on_submit=False):
    uploaded_files = st.file_uploader(
        "Drop PDF files here", 
        type="pdf", 
        accept_multiple_files=True
    )
    submit_button = st.form_submit_button(label="🚀 Convert to Excel")

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
                    if isinstance(value, str) and value.startswith("/"):
                        value = value.lstrip("/")
                    file_data[field_name] = value
                    
                all_form_data[uploaded_file.name] = file_data
                
            except Exception as e:
                st.error(f"Error reading {uploaded_file.name}: {e}")
            
            progress_bar.progress((index + 1) / len(uploaded_files))

        if all_form_data:
            df = pd.DataFrame.from_dict(all_form_data, orient="index")
            headers = [
                "Personal_name_first~7", "Personal_name_last~7", "EF_Emp_phone_primary~7", 
                "EF_Emp_birth_date~7", "EF_Emp_Residence_street_1~7", "EF_Emp_Residence_city~7", 
                "EF_Emp_Residence_state_rdo~7", "EF_Emp_Residence_zip_code~7", "EF_Emp_ssn~7", "EF_Emp_email~7"
            ]
            df = df.reindex(columns=headers)
            df.columns = ["First name", "Last name", "Phone", "DOB", "Address Street", "Address City", "Address State", "Address Zip", "SSN", "Email Address"]
            
            st.success(f"🎉 Successfully processed {len(all_form_data)} PDFs!")
            
            # --- NEW: VISUAL PREVIEW SECTION ---
            st.subheader("📊 Spreadsheet Preview")
            st.write("Review the extracted data below before downloading:")
            
            # Turns the preview into a grid you can highlight, copy, and select text from
            st.data_editor(df, use_container_width=True, key="excel_preview", disabled=True)
            
            # Prepare the Excel file in the background
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=True, sheet_name='Extracted Data')
            
            # Add some spacing and the download button right beneath the preview
            st.markdown("---")
            st.download_button(
                label="📥 Download Excel File",
                data=buffer.getvalue(),
                file_name="background_checks_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary" # Makes the button stand out in a bright color
            )
        else:
            st.warning("No interactive form fields were found in the uploaded PDFs.")

