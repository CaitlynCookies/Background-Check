import streamlit as st
import json
import pandas as pd
from pypdf import PdfReader
import io
import re
import difflib
import logging

# Silence the harmless Streamlit thread context warning in your console
logging.getLogger("streamlit.runtime.scriptrunner_utils").setLevel(logging.ERROR)

KEY_TRANSLATION_MAP = {
    # First Name
    "Personal_name_first~7": "First name",
    "Personal_name_first~8": "First name",
    "Personal_name_first~9": "First name",
    "EF_Emp_name_first~4": "First name",
    
    # Last Name
    "Personal_name_last~7": "Last name",
    "Personal_name_last~8": "Last name",
    "Personal_name_last~9": "Last name",
    "EF_Emp_name_last~4": "Last name",
    
    # Phone (Fixes the page ~3 and ~4 issue!)
    "EF_Emp_phone_primary~3": "Phone",
    "EF_Emp_phone_primary~4": "Phone",
    
    # DOB (Fixes the page ~4 issue!)
    "EF_Emp_birth_date~4": "DOB",
    
    # Street Address
    "ResAddrTran_addr__street_full_VC~7": "Address Street",
    "ResAddr_addr__street_1~8": "Address Street",
    "ResAddr_addr__street_1~9": "Address Street",
    "EF_Emp_Residence_street_1~3": "Address Street",
    "EF_Emp_Residence_street_1~4": "Address Street",
    
    # City
    "ResAddrTran_addr__city_cs_state_s_zip_VC~7": "Address City",
    "ResAddr_addr__city~8": "Address City",
    "ResAddr_addr__city~9": "Address City",
    "EF_Emp_Residence_city~3": "Address City",
    "EF_Emp_Residence_city~4": "Address City",
    
    # State
    "ResAddr_addr__state_desc~8": "Address State",
    "ResAddr_addr__state_desc~9": "Address State",
    "EF_Emp_Residence_state_rdo~3": "Address State",
    "EF_Emp_Residence_state_rdo~4": "Address State",
    
    # Zip Code
    "ResAddr_addr__zip_code~8": "Address Zip",
    "ResAddr_addr__zip_code~9": "Address Zip",
    "EF_Emp_Residence_zip_code~3": "Address Zip",
    "EF_Emp_Residence_zip_code~4": "Address Zip",
    
    # SSN
    "Personal_ssn~7": "SSN",
    "Personal_ssn~8": "SSN",
    "Personal_ssn~9": "SSN",
    "EF_Emp_ssn~4": "SSN",
    
    # Email (Fixes the page ~4 issue!)
    "EF_Emp_email~4": "Email Address"
}

# The precise column layout and order you want in your finalized Excel file
DESIRED_COLUMNS = [
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

def universal_standardize(raw_key: str) -> str:
    """
    Takes any messy PDF field key, strips page designations, 
    and cleanly maps it to a standard column name using fuzzy matching fallback.
    """
    # Step 1: Remove trailing page markers (e.g., 'Name~7' or 'SSN~10' -> 'Name' or 'SSN')
    clean_key = re.sub(r'~\d+(\.\d+)?(_\d+)?$', '', raw_key)
    
    # Step 2: Normalize (lowercase, strip whitespace)
    normalized_key = clean_key.strip().lower()
    
    # Step 3: Direct lookup in our simplified KEY_TRANSLATION_MAP
    if normalized_key in KEY_TRANSLATION_MAP:
        return KEY_TRANSLATION_MAP[normalized_key]
        
    # Step 4: Fuzzy Matching Fallback (85% match or closer) to catch minor typos automatically
    close_matches = difflib.get_close_matches(normalized_key, KEY_TRANSLATION_MAP.keys(), n=1, cutoff=0.85)
    if close_matches:
        matched_key = close_matches[0]
        return KEY_TRANSLATION_MAP[matched_key]
        
    # Step 5: If no match found, keep the stripped name as a fallback
    return clean_key


st.set_page_config(page_title="PDF to Excel Converter", page_icon="📝", layout="centered")

st.title("📝 PDF Background Check Extractor")
st.write("Upload your PDF forms below. The system automatically unifies differently labeled fields across form versions!")

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
                    
                    # Run Universal Standardizer
                    clean_header = universal_standardize(field_name)
                    
                    # --- CRITICAL OVERWRITE SAFEGUARD ---
                    # Only map or alter a value if the extracted string is non-empty.
                    # This prevents empty trailing forms in a packet from deleting valid phone/email/DOB fields.
                    if value:
                        if clean_header in file_data and file_data[clean_header]:
                            # If a value already exists, keep the longer, more complete value
                            if len(str(value)) > len(str(file_data[clean_header])):
                                file_data[clean_header] = value
                        else:
                            file_data[clean_header] = value
                    else:
                        # If incoming field is blank, only initialize it if the column isn't already populated
                        if clean_header not in file_data:
                            file_data[clean_header] = ""
                    
                all_form_data[uploaded_file.name] = file_data
                
            except Exception as e:
                st.error(f"Error reading {uploaded_file.name}: {e}")
            
            progress_bar.progress((index + 1) / len(uploaded_files))

        if all_form_data:
            df = pd.DataFrame.from_dict(all_form_data, orient="index")
            
            # Safety check: If the dataframe is completely empty, stop execution cleanly
            if df.empty:
                st.warning("Processed files, but no matching fields could be built.")
                st.stop()

            # Safely conform columns to your target headers in exact order (creates empty columns if missing)
            df = df.reindex(columns=DESIRED_COLUMNS)
            df.index.name = "Source File Name"
            
            # Clean up NaN representations to blanks for clear reading
            df = df.fillna("")
            
            st.success(f"🎉 Successfully processed {len(all_form_data)} PDFs!")
            
            # --- VISUAL PREVIEW SECTION ---
            st.subheader("📊 Spreadsheet Preview")
            st.write("Review the extracted data below before downloading:")
            st.data_editor(df, use_container_width=True, key="excel_preview", disabled=True)
            
            # Save final data into a clean memory buffer for browser-side download
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=True, sheet_name='Extracted Data')
            
            # Memory safety: clean up internal data frames to keep app connections stable
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
