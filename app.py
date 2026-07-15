import streamlit as st
import json
import pandas as pd
from pypdf import PdfReader
import io
import re

# ==========================================
# 1. VISUAL LAYER SCRAPER HELPER FUNCTION
# ==========================================
def extract_visual_fallbacks(reader):
    """
    Scrapes the visual text layer of the PDF as a fallback 
    for flattened, unreadable fields (DOB, Phone, Email).
    """
    fallback_data = {}
    
    # Combine the raw text from all pages
    full_text = ""
    for page in reader.pages:
        text = page.extract_text()
        if text:
            full_text += "\n" + text
            
    # Regex search for Email (looks for standard name@domain.com)
    email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', full_text)
    if email_match:
        fallback_data["Email Address"] = email_match.group(0)
        
    # Regex search for Phone Number (Requires separators like dashes, dots, or parentheses)
    phone_match = re.search(r'\(?\d{3}\)?[-.\s]\d{3}[-.\s]\d{4}', full_text)
    if phone_match:
        fallback_data["Phone"] = phone_match.group(0)
        
    # Regex search for DOB (Date of Birth)
    dob_label_match = re.search(
        r'(?:dob|date of birth|birth date|birthdate)[\s:]*(\d{1,2}[-/\s]\d{1,2}[-/\s]\d{2,4})', 
        full_text, 
        re.IGNORECASE
    )
    if dob_label_match:
        fallback_data["DOB"] = dob_label_match.group(1)
    else:
        # Generic date search fallback (MM/DD/YYYY or YYYY-MM-DD)
        date_match = re.search(r'\b(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})\b', full_text)
        if date_match:
            fallback_data["DOB"] = date_match.group(1)
            
    return fallback_data


# ==========================================
# 2. FIELD MAPPING & FINAL COLUMNS
# ==========================================
BASE_FIELD_MAP = {
    # --- FIRST NAME ---
    "First name": "First name",
    "Personal_name_first": "First name",
    "EF_Emp_name_first": "First name",
    
    # --- LAST NAME ---
    "Last name": "Last name",
    "Personal_name_last": "Last name",
    "EF_Emp_name_last": "Last name",
    
    # --- PHONE ---
    "Phone": "Phone",
    "EF_Emp_phone_primary": "Phone",
    "EF_Emp_phone_secondary": "Phone",
    
    # --- DOB ---
    "DOB": "DOB",
    "EF_Emp_birth_date": "DOB",
    
    # --- ADDRESS STREET ---
    "Address Street": "Address Street",
    "ResAddr_addr__street_1": "Address Street",
    "ResAddrTran_addr__street_full_VC": "Address Street",
    "EF_Emp_Residence_street_1": "Address Street",
    
    # --- ADDRESS CITY ---
    "Address City": "Address City",
    "ResAddr_addr__city": "Address City",
    "EF_Emp_Residence_city": "Address City",
    
    # --- ADDRESS STATE ---
    "Address State": "Address State",
    "ResAddr_addr__state_desc": "Address State",
    "EF_Emp_Residence_state_rdo": "Address State",
    
    # --- ADDRESS ZIP ---
    "Address Zip": "Address Zip",
    "ResAddr_addr__zip_code": "Address Zip",
    "EF_Emp_Residence_zip_code": "Address Zip",
    
    # --- SSN ---
    "SSN": "SSN",
    "Personal_ssn": "SSN",
    "EF_Emp_ssn": "SSN",
    
    # --- EMAIL ---
    "Email Address": "Email Address",
    "EF_Emp_email": "Email Address",
}

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

# ==========================================
# 3. STREAMLIT UI SETUP
# ==========================================
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

# ==========================================
# 4. PROCESSING ENGINE
# ==========================================
if submit_button:
    if not uploaded_files:
        st.error("Please upload at least one PDF file before clicking convert.")
    else:
        all_form_data = {}
        progress_bar = st.progress(0)
        
        for index, uploaded_file in enumerate(uploaded_files):
            try:
                reader = PdfReader(uploaded_file)
                file_data = {}
                
                # Standard Form Extraction
                fields = reader.get_fields() or {}
                text_fields = reader.get_form_text_fields() or {}
                combined_fields = {}
                
                # Process fields
                for field_name, field_info in fields.items():
                    val = field_info.get("/V", "")
                    if not val and "/DV" in field_info:
                        val = field_info.get("/DV", "")
                    combined_fields[field_name] = val
                
                for field_name, val in text_fields.items():
                    if val and not combined_fields.get(field_name):
                        combined_fields[field_name] = val

                # Clean & Map Fields to Base Names
                for field_name, value in combined_fields.items():
                    if isinstance(value, str):
                        value = value.strip()
                        if value.startswith("/"):
                            value = value.lstrip("/")
                    
                    base_field_name = field_name.split('~')[0] if '~' in field_name else field_name
                    clean_column_name = BASE_FIELD_MAP.get(base_field_name, field_name)
                    
                    if value:
                        if clean_column_name in file_data and file_data[clean_column_name]:
                            if len(str(value)) > len(str(file_data[clean_column_name])):
                                file_data[clean_column_name] = value
                        else:
                            file_data[clean_column_name] = value
                    else:
                        if clean_column_name not in file_data:
                            file_data[clean_column_name] = ""
                            
                # ==========================================
                # 5. THE SAFE FALLBACK TRIGGER (REVISED)
                # ==========================================
                # Only fallback if the fields are completely missing from the PDF's form data
                missing_crucial_fields = [
                    f for f in ["DOB", "Phone", "Email Address"] 
                    if not file_data.get(f) or file_data.get(f) == ""
                ]
                
                if missing_crucial_fields:
                    fallback_values = extract_visual_fallbacks(reader)
                    
                    for field in missing_crucial_fields:
                        if field in fallback_values and fallback_values[field]:
                            file_data[field] = fallback_values[field]
                            
                all_form_data[uploaded_file.name] = file_data
                
            except Exception as e:
                st.error(f"Error reading {uploaded_file.name}: {e}")
            
            progress_bar.progress((index + 1) / len(uploaded_files))

        # ==========================================
        # 6. OUTPUT GENERATION & DOWNLOADING
        # ==========================================
        if all_form_data:
            try:
                with open("extracted_data.json", "w", encoding="utf-8") as f:
                    json.dump(all_form_data, f, indent=4)
            except Exception:
                pass

            df = pd.DataFrame.from_dict(all_form_data, orient="index")
            df = df.reindex(columns=FINAL_COLUMNS)
            df.index.name = "Source File Name"
            df = df.fillna("")
            
            st.success(f"🎉 Successfully processed {len(all_form_data)} PDFs!")
            
            st.subheader("📊 Spreadsheet Preview")
            st.data_editor(df, use_container_width=True, key="excel_preview", disabled=True)
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, index=True, sheet_name='Extracted Data')
            
            st.markdown("---")
            
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
                st.download_button(
                    label="📥 Download Raw JSON File (Entire PDF Data)",
                    data=json.dumps(all_form_data, indent=4),
                    file_name="background_checks_complete_dump.json",
                    mime="application/json"
                )
        else:
            st.warning("No interactive form fields were found in the uploaded PDFs.")
