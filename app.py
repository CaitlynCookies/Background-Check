import streamlit as st
import json
import pandas as pd
from pypdf import PdfReader
import io

FIELD_MAP = {
    # === PERSONAL INFORMATION ===
    "Personal_name_first~7": "First Name",
    "Personal_name_first~8": "First Name",
    "Personal_name_first~9": "First Name",
    "EF_Emp_name_first~4": "First Name",
    "Personal_name_last~7": "Last Name",
    "Personal_name_last~8": "Last Name",
    "Personal_name_last~9": "Last Name",
    "EF_Emp_name_last~4": "Last Name",
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
    "EF_Emp_birth_date~4": "Date of Birth",
    "Personal_primary_language_english_yn~10": "Primary Language English (Y/N)",
    "Personal_primary_language_other_txt~10": "Primary Language Other",
    "Ethnicity_disabled_yn~4": "Ethnicity Disabled (Y/N)",
    
    # === CONTACT & ADDRESS ===
    "ResAddrTran_addr__street_full_VC~7": "Street Address",
    "ResAddr_addr__street_1~8": "Street Address 1",
    "ResAddr_addr__street_1~9": "Street Address 1",
    "EF_Emp_Residence_street_1~3": "Street Address 1",
    "EF_Emp_Residence_street_1~4": "Street Address 1",
    "ResAddr_addr__street_2~8": "Street Address 2",
    "ResAddr_addr__street_2~9": "Street Address 2",
    "EF_Emp_Residence_street_2~3": "Street Address 2",
    "ResAddrTran_addr__city_cs_state_s_zip_VC~7": "City State Zip",
    "ResAddr_addr__city~8": "City",
    "ResAddr_addr__city~9": "City",
    "EF_Emp_Residence_city~3": "City",
    "EF_Emp_Residence_city~4": "City",
    "ResAddr_addr__state_desc~8": "State",
    "ResAddr_addr__state_desc~9": "State",
    "EF_Emp_Residence_state_rdo~3": "State",
    "EF_Emp_Residence_state_rdo~4": "State",
    "ResAddr_addr__zip_code~8": "Zip Code",
    "ResAddr_addr__zip_code~9": "Zip Code",
    "EF_Emp_Residence_zip_code~3": "Zip Code",
    "EF_Emp_Residence_zip_code~4": "Zip Code",
    "EF_Emp_phone_primary~3": "Primary Phone",
    "EF_Emp_phone_primary~4": "Primary Phone",
    "EF_Emp_email~4": "Email Address",

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
    
    # === FEDERAL W-4 WORKSHEET DETAILS ===
    "Employee_W4_TwoJobs_line_1_amnt~7": "W4 Two Jobs Line 1 Amount",
    "Employee_W4_ThreeJobs_line_2a_amnt~7": "W4 Three Jobs Line 2a Amount",
    "Employee_W4_ThreeJobs_line_2b_amnt~7": "W4 Three Jobs Line 2b Amount",
    "Employee_W4_ThreeJobs_line_2c_amnt~7": "W4 Three Jobs Line 2c Amount",
    "Employee_W4_MultiJobs_line_4_amnt~7": "W4 Multi Jobs Line 4 Amount",
    "Employee_W4_MultiJobs_line_3_txt~7": "W4 Multi Jobs Line 3 Text",
    "Employee_W4_Deductions_wh_line_1_amnt~7": "W4 Deductions Line 1 Amount",
    "Employee_W4_Deductions_wh_line_2_amnt~7": "W4 Deductions Line 2 Amount",
    "Employee_W4_Deductions_wh_line_3_amnt~7": "W4 Deductions Line 3 Amount",
    "Employee_W4_Deductions_wh_line_4_amnt~7": "W4 Deductions Line 4 Amount",
    "Employee_W4_Deductions_wh_line_5_amnt~7": "W4 Deductions Line 5 Amount",

    # === NY STATE IT-2104 WITHHOLDINGS ===
    "W4_STATE~8": "State",
    "NewYorkIT2104_marital_status_W4_rdo~8": "NY Marital Status",
    "NewYorkIT2104_wh_resident_yn~8": "NY Resident (Y/N)",
    "NewYorkIT2104_wh_resident_yn~9": "NY Resident (Y/N)",
    "NewYorkIT2104~9.1_nonresident_yn": "NY Nonresident (Y/N)",
    "NewYorkIT2104_z_wh4_newyork_yonkers_yn~8": "Yonkers Resident (Y/N)",
    "NewYorkIT2104_z_wh4_newyork_yonkers_yn~9": "Yonkers Resident (Y/N)",
    "percent_estimate_NYS_yn~9": "Estimate NYS Tax (Y/N)",
    "percent_estimate_Yonkers_yn~9": "Estimate Yonkers Tax (Y/N)",
    "NewYorkIT2104~9.1_NTS_percent_txt": "NYS Tax Percentage",
    "NewYorkIT2104~9.1_Yonkers_percent_txt": "Yonkers Tax Percentage",
    "StateWithholding_wh_exemption_excesive_chk~8": "Excessive Exemption Check",
    "State Withholding_NEW_HIRE_YN~8": "New Hire State Withholding (Y/N)",

    # === NY IT-2104 WORKSHEET LINES ===
    "NewYorkIT2104_wh_line_1_amnt~8": "NY IT-2104 Line 1 Allowance",
    "NewYorkIT2104_wh_line_1_txt~8": "NY IT-2104 Line 1 Text",
    "NewYorkIT2104_wh_line_2_txt~8": "NY IT-2104 Line 2 Text",
    "NewYorkIT2104_wh_line_3_amnt~8": "NY IT-2104 Line 3 Allowance",
    "NewYorkIT2104_wh_line_4_amnt~8": "NY IT-2104 Line 4 Allowance",
    "NewYorkIT2104_wh_line_5_amnt~8": "NY IT-2104 Line 5 Allowance",
    "NewYorkIT2104_wh_line_6_txt~8": "NY IT-2104 Line 6 Text",
    "P_W4_NY_wh_line_6_txt~8": "NY IT-2104 Line 6 Text (Alt)",
    "NewYorkIT2104_wh_line_7_txt~8": "NY IT-2104 Line 7 Text",
    "NewYorkIT2104_wh_line_8_txt~8": "NY IT-2104 Line 8 Text",
    "NewYorkIT2104_wh_line_9_txt~8": "NY IT-2104 Line 9 Text",
    "NewYorkIT2104_wh_line_10_txt~8": "NY IT-2104 Line 10 Text",
    "NewYorkIT2104_wh_line_11_txt~8": "NY IT-2104 Line 11 Text",
    "NewYorkIT2104_wh_line_12_txt~8": "NY IT-2104 Line 12 Text",
    "NewYorkIT2104_wh_line_13_txt~8": "NY IT-2104 Line 13 Text",
    "NewYorkIT2104_wh_line_14_txt~8": "NY IT-2104 Line 14 Text",
    "NewYorkIT2104_wh_line_15_amnt~8": "NY IT-2104 Line 15 Amount",
    "NewYorkIT2104_wh_line_16_amnt~8": "NY IT-2104 Line 16 Amount",
    "P_W4_NY_wh_line_16_total~8": "NY IT-2104 Line 16 Total",
    "NewYorkIT2104_wh_line_17_txt~8": "NY IT-2104 Line 17 Text",
    "NewYorkIT2104_wh_line_19_amnt~8": "NY IT-2104 Line 19 Amount",
    "NewYorkIT2104_wh_line_21_amnt~8": "NY IT-2104 Line 21 Amount",
    "NewYorkIT2104_wh_line_22_amnt~8": "NY IT-2104 Line 22 Amount",
    "NewYorkIT2104_wh_line_23_amnt~8": "NY IT-2104 Line 23 Amount",
    "NewYorkIT2104_wh_line_25_txt~8": "NY IT-2104 Line 25 Text",
    "NewYorkIT2104_wh_line_26_amnt~8": "NY IT-2104 Line 26 Amount",
    "NewYorkIT2104_wh_line_27_amnt~8": "NY IT-2104 Line 27 Amount",
    "NewYorkIT2104_wh_line_28_amnt~8": "NY IT-2104 Line 28 Amount",
    "P_W4_NY_wh_line_30_txt~8": "NY IT-2104 Line 30 Text",
    "NewYorkIT2104_wh_line_34_txt~8": "NY IT-2104 Line 34 Text",

    # === NEW YORK WAGE NOTICE (WTPA / SECTION 195) ===
    "HR_ny_notice_given~10": "NY Notice Type Given",
    "HR_ny_pay_rate~10": "NY Regular Pay Rate",
    "HR_ny_ot_payrate~10": "NY Overtime Pay Rate",
    "HR_ny_pay_frequency~10": "NY Pay Frequency",
    "HR_ny_regular_payday~10": "NY Regular Payday",
    "HR_ny_allowances_none_yn~10": "NY Allowances - None (Y/N)",
    "HR_ny_allowances_tips_yn~10": "NY Allowances - Tips (Y/N)",
    "HR_ny_allowances_tips_amount~10": "NY Allowances - Tips Amount",
    "HR_ny_allowances_meals_yn~10": "NY Allowances - Meals (Y/N)",
    "HR_ny_allowances_meals_amount~10": "NY Allowances - Meals Amount",
    "HR_ny_allowances_lodging_yn~10": "NY Allowances - Lodging (Y/N)",
    "HR_ny_allowances_lodging_amount~10": "NY Allowances - Lodging Amount",
    "HR_ny_allowances_other_yn~10": "NY Allowances - Other (Y/N)",
    "HR_ny_allowances_other_txt~10": "NY Allowances - Other Text",

    # === COMPANY / EMPLOYER CONFIGURATION ===
    "W4Config_company_name~7": "Company Name",
    "W4Config_company_name~8": "Company Name",
    "W4Config_company_name~9": "Company Name",
    "W4Config_company_name~10": "Company Name",
    "ConfigNYEmp_dba_name~10": "Company DBA Name",
    "ConfigW4_ein_tax_id~7": "Employer EIN/Tax ID",
    "W4Config_ein_tax_id~8": "Employer EIN/Tax ID",
    "ConfigW4_ein_tax_id~10": "Employer EIN/Tax ID",
    "W4Config_Transforms_addr__full_s1_s2_c_st_zip_vc~7": "Company Address",
    "W4ConfigTran_addr_full_s1_s2_c_st_zip_VC~8": "Company Address",
    "W4Config_Transforms_addr__full_s1_s2_c_st_zip_vc~10": "Company Address",
    "W4ConfigTran_addr__street_full_VC~9": "Company Street Address",
    "W4Config_addr__city~9": "Company City",
    "W4Config_addr__state_desc~9": "Company State",
    "W4Config_addr__zip_code~9": "Company Zip",
    "Company_phone_work~10": "Company Work Phone",

    # === SIGNATURES & REPS ===
    "employee_authorization_signature~7": "Employee Signature (Y/N)",
    "employee_authorization_signature~8": "Employee Signature (Y/N)",
    "employee_authorization_signature~9": "Employee Signature (Y/N)",
    "employee_authorization_signature~10": "Employee Signature (Y/N)",
    "employee_authorization_signature~3": "Employee Signature (Y/N)",
    "employee_authorization_signature~4": "Employee Signature (Y/N)",
    "rep_click_signature~3": "Representative Signature (Y/N)",
    "rep_click_signature~4": "Representative Signature (Y/N)",
    "rep_click_name~10": "Representative Printed Name",
    "rep_click_title~10": "Representative Title",

    # === CUSTOM SETUP QUESTIONS ===
    "EF_HR_location_custom_questions_txt_1~3": "Location Custom Question 1",
    "EF_HR_custom_questions_txt_1~4": "HR Custom Question 1",
    "EF_HR_custom_questions_txt_2~4": "HR Custom Question 2",
    "EF_HR_location_custom_questions_txt_4~4": "Location Custom Question 4",
    
    # === SYSTEM / DESIGNATOR METADATA ===
    "DNM_Form_Type_Designator__B~8": "Form Designator B",
    "DNM_Form_Type_Designator__B_2~8": "Form Designator B 2",
}

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

