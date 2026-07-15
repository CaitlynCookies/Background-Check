import streamlit as st
import json
import pandas as pd
from pypdf import PdfReader
import io
import re
import difflib
import logging


# 1. Simplified Field Map (No ~7, ~8, ~10 page suffixes needed anymore!)
FIELD_MAP = {
    "personal_name_first": "First Name",
    "ef_emp_name_first": "First Name",
    "personal_name_last": "Last Name",
    "ef_emp_name_last": "Last Name",
    "personaltransforms_name_middle_initial_vb": "Middle Initial",
    "personaltran_name_middle_init": "Middle Initial",
    "ef_emp_name_middle": "Middle Initial",
    "personaltransforms_name_full_vc": "Full Name",
    "ef_emp_name_full": "Full Name",
    "personaltran_name__full_f_s_m_s_l_vc": "Full Name",
    "personal_ssn": "SSN",
    "ef_emp_ssn": "SSN",
    "ef_emp_birth_date": "Date of Birth",
    "ef_emp_phone_primary": "Primary Phone",
    "ef_emp_email": "Email Address",
    "resaddrtran_addr__street_full_vc": "Street Address",
    "resaddr_addr__street_1": "Street Address 1",
    "ef_emp_residence_street_1": "Street Address 1",
    "resaddr_addr__street_2": "Street Address 2",
    "ef_emp_residence_street_2": "Street Address 2",
    "resaddr_addr__city": "City",
    "ef_emp_residence_city": "City",
    "resaddr_addr__state_desc": "State",
    "ef_emp_residence_state_rdo": "State",
    "resaddr_addr__zip_code": "Zip Code",
    "ef_emp_residence_zip_code": "Zip Code",
}

# Define the final columns we want in our Excel sheet (in order)
DESIRED_COLUMNS = [
    "First Name", 
    "Last Name", 
    "Primary Phone", 
    "Date of Birth", 
    "Street Address 1", 
    "City", 
    "State", 
    "Zip Code", 
    "SSN", 
    "Email Address"
]

def universal_standardize(raw_key: str) -> str:
    """
    Takes any messy PDF field key, strips page designations, 
    and cleanly maps it to a standard column name using fuzzy matching fallback.
    """
    # Step 1: Remove trailing page markers (e.g., 'Name~7' or 'SSN~10' -> 'Name' or 'SSN')
    clean_key = re.sub(r'~\d+(\.\d+)?(_\d+)?$', '', raw_key)
    
    # Step 2: Normalize (lowercase, strip whitespace)
    normalized_key = clean_key.strip().lower()
    
    # Step 3: Direct lookup in our simplified FIELD_MAP
    if normalized_key in FIELD_MAP:
        return FIELD_MAP[normalized_key]
        
    # Step 4: Fuzzy Matching Fallback (85% match or closer)
    close_matches = difflib.get_close_matches(normalized_key, FIELD_MAP.keys(), n=1, cutoff=0.85)
    if close_matches:
        matched_key = close_matches[0]
        return FIELD_MAP[matched_key]
        
    # Step 5: If no match found, keep the stripped name as a fallback
    return clean_key


st.set_page_config(page_title="PDF to Excel Converter", page_icon="📝", layout="centered")

st.title("📝 PDF Background Check Extractor")
st.write("Upload your PDF forms below. The system will automatically align varying field names from different versions!")

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
                    if isinstance(value, str) and value.startswith("/"):
                        value = value.lstrip("/")
                    
                    # Run Universal Standardizer
                    clean_header = universal_standardize(field_name)
                    
                    # Combine overlapping/duplicate field values intelligently
                    if clean_header in file_data and file_data[clean_header]:
                        if len(str(value)) > len(str(file_data[clean_header])):
                            file_data[clean_header] = value
                    else:
                        file_data[clean_header] = value
                    
                all_form_data[uploaded_file.name] = file_data
                
            except Exception as e:
                st.error(f"Error reading {uploaded_file.name}: {e}")
            
            progress_bar.progress((index + 1) / len(uploaded_files))

        if all_form_data:
            df = pd.DataFrame.from_dict(all_form_data, orient="index")
            
            # Safety check: If the dataframe is completely empty, stop processing safely
            if df.empty:
                st.warning("Processed files, but no matching columns could be built.")
                st.stop()

            # Reindex to keep only our beautifully organized target columns
            df = df.reindex(columns=DESIRED_COLUMNS)
            df.index.name = "Source File Name"
            
            # Clean up missing data visually in Streamlit (replaces NaN with blanks)
            df = df.fillna("")
            
            st.success(f"🎉 Successfully processed {len(all_form_data)} PDFs!")
            
            st.subheader("📊 Spreadsheet Preview")
            st.data_editor(df, use_container_width=True, key="excel_preview", disabled=True)
            
            # Write to a memory buffer for the browser to download
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=True, sheet_name='Extracted Data')
            
            # Explicitly delete objects to free system memory and prevent dropped connections
            del df
            del all_form_data

            st.markdown("---")
            st.download_button(
                label="📥 Download Excel File",
                data=buffer.getvalue(),
                file_name="background_checks_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        else:
            st.warning("No interactive form fields were found in the uploaded PDFs.")
