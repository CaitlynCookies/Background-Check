import streamlit as st
import json
import pandas as pd
from pypdf import PdfReader
import io

FIELD_MAP = {
    # === PERSONAL INFORMATION ===
    "Personal_name_first~7": "First name",
    "Personal_name_first~8": "First name",
    "Personal_name_first~9": "First name",
    "EF_Emp_name_first~4": "First name",
    "Personal_name_last~7": "Last name",
    "Personal_name_last~8": "Last name",
    "Personal_name_last~9": "Last name",
    "EF_Emp_name_last~4": "Last name",
    "PersonalTransforms_name_middle_initial_VB~7": "Middle Initial",
    "PersonalTran_name_middle_init~8": "Middle Initial",
    "PersonalTran_name_middle_init~9": "Middle Initial",
    "EF_Emp_name_middle~4": "Middle Initial",
    "PersonalTransforms_name_full_VC~4": "Full Name",
    "EF_Emp_name_full~3": "Full Name",
    "EF_Emp_name_full~4": "Full Name",
    "PersonalTran_name__full_f_s_m_s_l_vc~10": "Full Name",
    "Personal_ssn~7": "SSN",
    "Personal_ssn~8": "SSN",
    "Personal_ssn~9": "SSN",
    "EF_Emp_ssn~4": "SSN",
    
    # --- FIXED DATE OF BIRTH ---
    "EF_Emp_birth_date~4": "DOB",
    "EF_Emp_birth_date~7": "DOB",
    
    "Personal_primary_language_english_yn~10": "Primary Language English (Y/N)",
    "Personal_primary_language_other_txt~10": "Primary Language Other",
    "Ethnicity_disabled_yn~4": "Ethnicity Disabled (Y/N)",
    
    # === CONTACT & ADDRESS ===
    "ResAddrTran_addr__street_full_VC~7": "Address Street",
    "ResAddr_addr__street_1~8": "Address Street",
    "ResAddr_addr__street_1~9": "Address Street",
    "EF_Emp_Residence_street_1~3": "Address Street",
    "EF_Emp_Residence_street_1~4": "Address Street",
    "ResAddr_addr__street_2~8": "Address Street 2",
    "ResAddr_addr__street_2~9": "Address Street 2",
    "EF_Emp_Residence_street_2~3": "Address Street 2",
    "ResAddrTran_addr__city_cs_state_s_zip_VC~7": "City State Zip",
    "ResAddr_addr__city~8": "Address City",
    "ResAddr_addr__city~9": "Address City",
    "EF_Emp_Residence_city~3": "Address City",
    "EF_Emp_Residence_city~4": "Address City",
    "ResAddr_addr__state_desc~8": "Address State",
    "ResAddr_addr__state_desc~9": "Address State",
    "EF_Emp_Residence_state_rdo~3": "Address State",
    "EF_Emp_Residence_state_rdo~4": "Address State",
    "ResAddr_addr__zip_code~8": "Address Zip",
    "ResAddr_addr__zip_code~9": "Address Zip",
    "EF_Emp_Residence_zip_code~3": "Address Zip",
    "EF_Emp_Residence_zip_code~4": "Address Zip",
    
    # --- FIXED PHONE ---
    "EF_Emp_phone_primary~3": "Phone",
    "EF_Emp_phone_primary~4": "Phone",
    "EF_Emp_phone_primary~7": "Phone",
    
    # --- FIXED EMAIL ---
    "EF_Emp_email~4": "Email Address",
    "EF_Emp_email~7": "Email Address",

    # === EMPLOYMENT & HR DETAILS ===
    "EmployeeID~4": "Employee ID",
    "AS_z_wfuel_oracle_id~8": "Oracle ID",
    "AS_z_wfuel_oracle_id_2~8": "Oracle ID 2",
    "HR_Job_Title~4": "Job Title",
    "HR_date_of_hire~4": "Date of Hire",
    "HR_Profile_date_start~7": "Start Date",
    "HR_date_start~8": "Start Date",
    "EF_Emp_start_date~4": "Start Date",
    "EF_Emp_pay_rate~4": "Pay Rate",
    "EF_Emp_hours_per_week~4": "Hours Per Week",
    "The_Date~4": "Form Date",

    # === FEDERAL W-4 WITHHOLDINGS ===
    "W4_marital_status_W4_rdo~7": "Federal Marital Status",
    "W4_two_job_method_rdo~7": "Federal Two Jobs Option (Y/N)",
    "Employee_W4_extra_withhold_amount~7": "W4 Extra Withholding Amount",
    "Employee_W4_deductions_amount~7": "W4 Deductions Amount",
    "Employee_W4_other_income_amount~7": "W4 Other Income Amount",
    "Employee_W4_qualifying_deps_amount~7": "W4 Qualifying Dependents Amount",
    "Employee_W4_other_deps_amount~7": "W4 Other Dependents Amount",
    "Employee_W4_total_deps_amount~7": "W4 Total Dependents Amount",
    "Employee_W4_Exempt_txt~7": "W4 Exempt Status Text",
    "Employee_W4_NRA_txt~7": "W4 Non-Resident Alien Text",
}

# The final friendly column names in the exact order you want them
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
                    
                    # Convert raw PDF key to final standardized Excel column name
                    # If it's not in FIELD_MAP, keep the original name
                    clean_column_name = FIELD_MAP.get(field_name, field_name)
                    
                    # Only map if the value is non-empty (prevents blank fields from overwriting valid ones)
                    if value:
                        if clean_column_name in file_data and file_data[clean_column_name]:
                            # Keep the longer value if there is a conflict
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
            df = pd.DataFrame.from_dict(all_form_data, orient="index")
            
            # Reindex to our clean friendly column headers
            df = df.reindex(columns=FINAL_COLUMNS)
            df.index.name = "Source File Name"
            
            # Replaces missing values with clean blank spaces
            df = df.fillna("")
            
            st.success(f"🎉 Successfully processed {len(all_form_data)} PDFs!")
            
            # --- VISUAL PREVIEW SECTION ---
            st.subheader("📊 Spreadsheet Preview")
            st.write("Review the extracted data below before downloading:")
            
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
                type="primary"
            )
        else:
            st.warning("No interactive form fields were found in the uploaded PDFs.")
