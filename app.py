import streamlit as st
import json
import pandas as pd
from pypdf import PdfReader
import io

# This map explicitly targets the raw technical keys from your JSON output 
# and turns them into readable column headers for your Excel file.
FIELD_MAP = {
    # --- PAGE 3 ---
    "EF_Emp_Residence_street_1~3": "Address Street",
    "EF_Emp_Residence_street_2~3": "Address Street 2",
    "EF_Emp_Residence_city~3": "Address City",
    "EF_Emp_Residence_state_rdo~3": "Address State",
    "EF_Emp_Residence_zip_code~3": "Address Zip",
    "employee_authorization_signature~3": "Employee Signature (Pg 3)",
    "rep_click_signature~3": "Rep Signature (Pg 3)",
    "EF_HR_location_custom_questions_txt_1~3": "HR Location Question 1 (Pg 3)",

    # --- PAGE 4 ---
    "employee_authorization_signature~4": "Employee Signature (Pg 4)",
    "rep_click_signature~4": "Rep Signature (Pg 4)",
    "EF_HR_custom_questions_txt_1~4": "HR Custom Question 1 (Pg 4)",
    "EF_HR_custom_questions_txt_2~4": "HR Custom Question 2 (Pg 4)",
    "EF_HR_location_custom_questions_txt_4~4": "HR Location Question 4 (Pg 4)",

    # --- PAGE 7 (W-4) ---
    "employee_authorization_signature~7": "Employee Signature (Pg 7)",
    "W4Config_company_name~7": "W4 Company Name",
    "W4Config_Transforms_addr__full_s1_s2_c_st_zip_vc~7": "W4 Company Address",
    "ConfigW4_ein_tax_id~7": "W4 EIN/Tax ID",
    "Employee_W4_TwoJobs_line_1_amnt~7": "W4 Two Jobs Line 1",
    "Employee_W4_ThreeJobs_line_2a_amnt~7": "W4 Three Jobs Line 2a",
    "Employee_W4_ThreeJobs_line_2b_amnt~7": "W4 Three Jobs Line 2b",
    "Employee_W4_ThreeJobs_line_2c_amnt~7": "W4 Three Jobs Line 2c",
    "Employee_W4_MultiJobs_line_4_amnt~7": "W4 Multi Jobs Line 4",
    "Employee_W4_MultiJobs_line_3_txt~7": "W4 Multi Jobs Line 3 Text",
    "Employee_W4_Deductions_wh_line_1_amnt~7": "W4 Deductions Line 1",
    "Employee_W4_Deductions_wh_line_2_amnt~7": "W4 Deductions Line 2",
    "Employee_W4_Deductions_wh_line_3_amnt~7": "W4 Deductions Line 3",
    "Employee_W4_Deductions_wh_line_4_amnt~7": "W4 Deductions Line 4",
    "Employee_W4_Deductions_wh_line_5_amnt~7": "W4 Deductions Line 5",

    # --- PAGE 8 (NY IT-2104) ---
    "NewYorkIT2104_marital_status_W4_rdo~8": "NY Marital Status",
    "NewYorkIT2104_wh_resident_yn~8": "NY Resident (Y/N)",
    "NewYorkIT2104_z_wh4_newyork_yonkers_yn~8": "NY Yonkers Resident (Y/N)",
    "employee_authorization_signature~8": "Employee Signature (Pg 8)",
    "StateWithholding_wh_exemption_excesive_chk~8": "NY Excessive Withholding",
    "State Withholding_NEW_HIRE_YN~8": "NY New Hire (Y/N)",
    "W4Config_ein_tax_id~8": "NY EIN/Tax ID",
    "W4Config_company_name~8": "NY Company Name",
    "W4ConfigTran_addr_full_s1_s2_c_st_zip_VC~8": "NY Company Address",
    "DNM_Form_Type_Designator__B~8": "Form Designator B",
    "DNM_Form_Type_Designator__B_2~8": "Form Designator B 2",
    "W4_STATE~8": "W4 State",
    "NewYorkIT2104_wh_line_1_amnt~8": "NY Line 1 Amount",
    "NewYorkIT2104_wh_line_3_amnt~8": "NY Line 3 Amount",
    "NewYorkIT2104_wh_line_4_amnt~8": "NY Line 4 Amount",
    "NewYorkIT2104_wh_line_5_amnt~8": "NY Line 5 Amount",
    "NewYorkIT2104_wh_line_15_amnt~8": "NY Line 15 Amount",
    "NewYorkIT2104_wh_line_16_amnt~8": "NY Line 16 Amount",
    "NewYorkIT2104_wh_line_19_amnt~8": "NY Line 19 Amount",
    "NewYorkIT2104_wh_line_21_amnt~8": "NY Line 21 Amount",
    "NewYorkIT2104_wh_line_22_amnt~8": "NY Line 22 Amount",
    "NewYorkIT2104_wh_line_23_amnt~8": "NY Line 23 Amount",
    "NewYorkIT2104_wh_line_26_amnt~8": "NY Line 26 Amount",
    "NewYorkIT2104_wh_line_27_amnt~8": "NY Line 27 Amount",
    "NewYorkIT2104_wh_line_28_amnt~8": "NY Line 28 Amount",
    "P_W4_NY_wh_line_16_total~8": "NY Total Line 16",
    "NewYorkIT2104_wh_line_1_txt~8": "NY Line 1 Text",
    "NewYorkIT2104_wh_line_2_txt~8": "NY Line 2 Text",
    "NewYorkIT2104_wh_line_6_txt~8": "NY Line 6 Text",
    "NewYorkIT2104_wh_line_7_txt~8": "NY Line 7 Text",
    "NewYorkIT2104_wh_line_8_txt~8": "NY Line 8 Text",
    "NewYorkIT2104_wh_line_9_txt~8": "NY Line 9 Text",
    "NewYorkIT2104_wh_line_10_txt~8": "NY Line 10 Text",
    "NewYorkIT2104_wh_line_11_txt~8": "NY Line 11 Text",
    "NewYorkIT2104_wh_line_12_txt~8": "NY Line 12 Text",
    "NewYorkIT2104_wh_line_13_txt~8": "NY Line 13 Text",
    "NewYorkIT2104_wh_line_14_txt~8": "NY Line 14 Text",
    "NewYorkIT2104_wh_line_17_txt~8": "NY Line 17 Text",
    "NewYorkIT2104_wh_line_25_txt~8": "NY Line 25 Text",
    "NewYorkIT2104_wh_line_34_txt~8": "NY Line 34 Text",
    "P_W4_NY_wh_line_6_txt~8": "NY Line 6 P_W4 Text",
    "P_W4_NY_wh_line_30_txt~8": "NY Line 30 P_W4 Text",

    # --- PAGE 9 (NY IT-2104 Page 2) ---
    "employee_authorization_signature~9": "Employee Signature (Pg 9)",
    "NewYorkIT2104~9": "NY IT-2104 Pg 9",
    "NewYorkIT2104~9.1_nonresident_yn": "NY Non-Resident (Y/N)",
    "NewYorkIT2104~9.1_NTS_percent_txt": "NY NTS Percent",
    "NewYorkIT2104~9.1_Yonkers_percent_txt": "NY Yonkers Percent",
    "NewYorkIT2104_wh_resident_yn~9": "NY Resident Pg 9 (Y/N)",
    "NewYorkIT2104_z_wh4_newyork_yonkers_yn~9": "NY Yonkers Resident Pg 9 (Y/N)",
    "percent_estimate_NYS_yn~9": "NY Est Percent (Y/N)",
    "percent_estimate_Yonkers_yn~9": "Yonkers Est Percent (Y/N)",
    "W4Config_company_name~9": "Company Name Pg 9",
    "W4Config_addr__city~9": "Company City Pg 9",
    "W4ConfigTran_addr__street_full_VC~9": "Company Street Pg 9",
    "W4Config_addr__state_desc~9": "Company State Pg 9",
    "W4Config_addr__zip_code~9": "Company Zip Pg 9",

    # --- PAGE 10 (WT-240 / HR Details) ---
    "W4Config_company_name~10": "Company Name Pg 10",
    "ConfigW4_ein_tax_id~10": "EIN Pg 10",
    "W4Config_Transforms_addr__full_s1_s2_c_st_zip_vc~10": "Company Address Pg 10",
    "Company_phone_work~10": "Company Phone Pg 10",
    "ConfigNYEmp_dba_name~10": "Company DBA",
    "HR_ny_notice_given~10": "NY Notice Given",
    "HR_ny_pay_rate~10": "NY Pay Rate",
    "HR_ny_allowances_none_yn~10": "NY Allowances None (Y/N)",
    "HR_ny_allowances_tips_yn~10": "NY Allowances Tips (Y/N)",
    "HR_ny_allowances_meals_yn~10": "NY Allowances Meals (Y/N)",
    "HR_ny_allowances_lodging_yn~10": "NY Allowances Lodging (Y/N)",
    "HR_ny_allowances_other_yn~10": "NY Allowances Other (Y/N)",
    "HR_ny_allowances_tips_amount~10": "NY Tips Amount",
    "HR_ny_allowances_meals_amount~10": "NY Meals Amount",
    "HR_ny_allowances_lodging_amount~10": "NY Lodging Amount",
    "HR_ny_allowances_other_txt~10": "NY Other Allowances Text",
    "HR_ny_regular_payday~10": "NY Regular Payday",
    "HR_ny_pay_frequency~10": "NY Pay Frequency",
    "HR_ny_ot_payrate~10": "NY OT Payrate",
    "rep_click_title~10": "Rep Title",
    "rep_click_name~10": "Rep Name",
    "employee_authorization_signature~10": "Employee Signature (Pg 10)",
}

# The main columns we want to force to the very front of the Excel sheet
CORE_COLUMNS = [
    "First name", 
    "Middle Initial",
    "Last name", 
    "Phone", 
    "Email Address",
    "DOB", 
    "SSN", 
    "Address Street", 
    "Address Street 2",
    "Address City", 
    "Address State", 
    "Address Zip", 
    "City State Zip",
    "Full Name"
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
                    
                    # Convert raw PDF key to standardized Excel column name
                    clean_column_name = FIELD_MAP.get(field_name, field_name)
                    
                    # Overwrite protection: keeps the longest string to prevent 
                    # a blank field on page 10 from overwriting a filled field on page 3
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
            df = pd.DataFrame.from_dict(all_form_data, orient="index")
            df.index.name = "Source File Name"
            df = df.fillna("")
            
            # --- SMART COLUMN ORDERING ---
            # 1. Grab all columns that the PDF actually spit out
            all_extracted_cols = list(df.columns)
            
            # 2. Find which of our CORE_COLUMNS exist in the data
            present_core_cols = [col for col in CORE_COLUMNS if col in all_extracted_cols]
            
            # 3. Gather everything else that isn't a core column (so we don't lose the tax/HR data)
            other_cols = [col for col in all_extracted_cols if col not in present_core_cols]
            
            # 4. Rearrange the DataFrame: Core columns first, then everything else.
            # This ensures NO columns are dropped or go missing.
            df = df[present_core_cols + other_cols]
            
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
                    data=json.dumps(all_form_data, indent=4),
                    file_name="background_checks_output.json",
                    mime="application/json"
                )
        else:
            st.warning("No interactive form fields were found in the uploaded PDFs.")
