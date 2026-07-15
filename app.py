import streamlit as st
import json
import pandas as pd
from pypdf import PdfReader
import io

FIELD_MAP = {
    # === PERSONAL INFORMATION ===
    "First name": "First name",
    "Last name": "Last name",
    "Middle Initial": "Middle Initial",
    "Full Name": "Full Name",
    "SSN": "SSN",
    "DOB": "DOB",
    "Phone": "Phone",
    "Email Address": "Email Address",
    "Primary Language English (Y/N)": "Primary Language English (Y/N)",
    "Primary Language Other": "Primary Language Other",
    "Ethnicity Disabled (Y/N)": "Ethnicity Disabled (Y/N)",

    # === CONTACT & ADDRESS ===
    "Address Street": "Address Street",
    "Address Street 2": "Address Street 2",
    "Address City": "Address City",
    "Address State": "Address State",
    "Address Zip": "Address Zip",
    "City State Zip": "City State Zip",

    # === EMPLOYMENT & HR DETAILS ===
    "Employee ID": "Employee ID",
    "Oracle ID": "Oracle ID",
    "Oracle ID 2": "Oracle ID 2",
    "Job Title": "Job Title",
    "Date of Hire": "Date of Hire",
    "Start Date": "Start Date",
    "Pay Rate": "Pay Rate",
    "Hours Per Week": "Hours Per Week",
    "Form Date": "Form Date",

    # === FEDERAL W-4 WITHHOLDINGS ===
    "Federal Marital Status": "Federal Marital Status",
    "Federal Two Jobs Option (Y/N)": "Federal Two Jobs Option (Y/N)",
    "W4 Extra Withholding Amount": "W4 Extra Withholding Amount",
    "W4 Deductions Amount": "W4 Deductions Amount",
    "W4 Other Income Amount": "W4 Other Income Amount",
    "W4 Qualifying Dependents Amount": "W4 Qualifying Dependents Amount",
    "W4 Other Dependents Amount": "W4 Other Dependents Amount",
    "W4 Total Dependents Amount": "W4 Total Dependents Amount",
    "W4 Exempt Status Text": "W4 Exempt Status Text",
    "W4 Non-Resident Alien Text": "W4 Non-Resident Alien Text",
    
    "W4Config_company_name~7": "W4 Company Name",
    "W4Config_company_name~8": "W4 Company Name",
    "W4Config_company_name~9": "W4 Company Name",
    "W4Config_company_name~10": "W4 Company Name",
    
    "ConfigW4_ein_tax_id~7": "EIN/Tax ID",
    "W4Config_ein_tax_id~8": "EIN/Tax ID",
    "ConfigW4_ein_tax_id~10": "EIN/Tax ID",
    
    "W4Config_Transforms_addr__full_s1_s2_c_st_zip_vc~7": "W4 Company Address",
    "W4ConfigTran_addr_full_s1_s2_c_st_zip_VC~8": "W4 Company Address",
    "W4Config_Transforms_addr__full_s1_s2_c_st_zip_vc~10": "W4 Company Address",
    
    "W4ConfigTran_addr__street_full_VC~9": "W4 Company Street",
    "W4Config_addr__city~9": "W4 Company City",
    "W4Config_addr__state_desc~9": "W4 Company State",
    "W4Config_addr__zip_code~9": "W4 Company Zip",

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

    # === NEW YORK STATE TAX WITHHOLDINGS (IT-2104) ===
    "NewYorkIT2104_marital_status_W4_rdo~8": "NY Marital Status",
    "NewYorkIT2104_wh_resident_yn~8": "NY Resident (Y/N)",
    "NewYorkIT2104_wh_resident_yn~9": "NY Resident (Y/N)",
    "NewYorkIT2104_z_wh4_newyork_yonkers_yn~8": "NY Yonkers Resident (Y/N)",
    "NewYorkIT2104_z_wh4_newyork_yonkers_yn~9": "NY Yonkers Resident (Y/N)",
    "W4_STATE~8": "Withholding State",
    
    "NewYorkIT2104_wh_line_1_amnt~8": "NY Allowances Line 1",
    "NewYorkIT2104_wh_line_3_amnt~8": "NY Allowances Line 3",
    "NewYorkIT2104_wh_line_4_amnt~8": "NY Allowances Line 4",
    "NewYorkIT2104_wh_line_5_amnt~8": "NY Allowances Line 5",
    "NewYorkIT2104_wh_line_15_amnt~8": "NY Allowances Line 15",
    "NewYorkIT2104_wh_line_16_amnt~8": "NY Allowances Line 16",
    "NewYorkIT2104_wh_line_19_amnt~8": "NY Allowances Line 19",
    "NewYorkIT2104_wh_line_21_amnt~8": "NY Allowances Line 21",
    "NewYorkIT2104_wh_line_22_amnt~8": "NY Allowances Line 22",
    "NewYorkIT2104_wh_line_23_amnt~8": "NY Allowances Line 23",
    "NewYorkIT2104_wh_line_26_amnt~8": "NY Allowances Line 26",
    "NewYorkIT2104_wh_line_27_amnt~8": "NY Allowances Line 27",
    "NewYorkIT2104_wh_line_28_amnt~8": "NY Allowances Line 28",
    "P_W4_NY_wh_line_16_total~8": "NY Allowances Total Line 16",
    
    "NewYorkIT2104_wh_line_1_txt~8": "NY Withholding Text Line 1",
    "NewYorkIT2104_wh_line_2_txt~8": "NY Withholding Text Line 2",
    "NewYorkIT2104_wh_line_6_txt~8": "NY Withholding Text Line 6",
    "NewYorkIT2104_wh_line_7_txt~8": "NY Withholding Text Line 7",
    "NewYorkIT2104_wh_line_8_txt~8": "NY Withholding Text Line 8",
    "NewYorkIT2104_wh_line_9_txt~8": "NY Withholding Text Line 9",
    "NewYorkIT2104_wh_line_10_txt~8": "NY Withholding Text Line 10",
    "NewYorkIT2104_wh_line_11_txt~8": "NY Withholding Text Line 11",
    "NewYorkIT2104_wh_line_12_txt~8": "NY Withholding Text Line 12",
    "NewYorkIT2104_wh_line_13_txt~8": "NY Withholding Text Line 13",
    "NewYorkIT2104_wh_line_14_txt~8": "NY Withholding Text Line 14",
    "NewYorkIT2104_wh_line_17_txt~8": "NY Withholding Text Line 17",
    "NewYorkIT2104_wh_line_25_txt~8": "NY Withholding Text Line 25",
    "NewYorkIT2104_wh_line_34_txt~8": "NY Withholding Text Line 34",
    "P_W4_NY_wh_line_6_txt~8": "NY Withholding Text Line 6 (P_W4)",
    "P_W4_NY_wh_line_30_txt~8": "NY Withholding Text Line 30 (P_W4)",
    
    "NewYorkIT2104~9": "NY IT-2104 Page 9 Status",
    "NewYorkIT2104~9.1_nonresident_yn": "NY Non-Resident Status (Y/N)",
    "NewYorkIT2104~9.1_NTS_percent_txt": "NY Non-Resident NYS Work Percent",
    "NewYorkIT2104~9.1_Yonkers_percent_txt": "NY Non-Resident Yonkers Work Percent",
    "percent_estimate_NYS_yn~9": "Estimate NYS Tax (Y/N)",
    "percent_estimate_Yonkers_yn~9": "Estimate Yonkers Tax (Y/N)",
    "StateWithholding_wh_exemption_excesive_chk~8": "Excessive Withholding Exemption (Y/N)",
    "State Withholding_NEW_HIRE_YN~8": "New Hire (Y/N)",

    # === NY SPECIAL HR LABELS (WT-240) ===
    "ConfigNYEmp_dba_name~10": "Employer DBA Name",
    "Company_phone_work~10": "Employer Work Phone",
    "HR_ny_notice_given~10": "NY Notice Given Format",
    "HR_ny_pay_rate~10": "NY Pay Rate Amount",
    "HR_ny_allowances_none_yn~10": "NY Allowances None (Y/N)",
    "HR_ny_allowances_tips_yn~10": "NY Allowances Tips (Y/N)",
    "HR_ny_allowances_meals_yn~10": "NY Allowances Meals (Y/N)",
    "HR_ny_allowances_lodging_yn~10": "NY Allowances Lodging (Y/N)",
    "HR_ny_allowances_other_yn~10": "NY Allowances Other (Y/N)",
    "HR_ny_allowances_tips_amount~10": "NY Allowances Tips Amount",
    "HR_ny_allowances_meals_amount~10": "NY Allowances Meals Amount",
    "HR_ny_allowances_lodging_amount~10": "NY Allowances Lodging Amount",
    "HR_ny_allowances_other_txt~10": "NY Allowances Other Description",
    "HR_ny_regular_payday~10": "NY Regular Payday Schedule",
    "HR_ny_pay_frequency~10": "NY Pay Frequency",
    "HR_ny_ot_payrate~10": "NY Overtime Pay Rate",

    # === FORMS DESIGNATORS & CUSTOM HR QUESTIONS ===
    "DNM_Form_Type_Designator__B~8": "Form Designator B",
    "DNM_Form_Type_Designator__B_2~8": "Form Designator B 2",
    "EF_HR_location_custom_questions_txt_1~3": "Location Custom Question 1 (Pg 3)",
    "EF_HR_custom_questions_txt_1~4": "Custom Question 1 (Pg 4)",
    "EF_HR_custom_questions_txt_2~4": "Custom Question 2 (Pg 4)",
    "EF_HR_location_custom_questions_txt_4~4": "Location Custom Question 4 (Pg 4)",

    # === SIGNATURES & AUTHORIZATIONS ===
    "employee_authorization_signature~3": "Employee Signature (Pg 3)",
    "employee_authorization_signature~4": "Employee Signature (Pg 4)",
    "employee_authorization_signature~7": "Employee Signature (Pg 7)",
    "employee_authorization_signature~8": "Employee Signature (Pg 8)",
    "employee_authorization_signature~9": "Employee Signature (Pg 9)",
    "employee_authorization_signature~10": "Employee Signature (Pg 10)",
    "rep_click_signature~3": "Representative Signature (Pg 3)",
    "rep_click_signature~4": "Representative Signature (Pg 4)",
    "rep_click_name~10": "Authorized Representative Name",
    "rep_click_title~10": "Authorized Representative Title",
}

# Clean, structured column presentation order for the final spreadsheet output
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
                    
                    # Convert raw PDF key to standardized Excel column name
                    clean_column_name = FIELD_MAP.get(field_name, field_name)
                    
                    # Overwrite protection logic to prevent blanks on subsequent pages from erasing filled ones
                    if value:
                        if clean_column_name in file_data and file_data[clean_column_name]:
                            # Keep whichever string value is longer/more complete
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
            # === AUTO-SAVE JSON TO COLAB FILE SYSTEM ===
            try:
                with open("extracted_data.json", "w", encoding="utf-8") as f:
                    json.dump(all_form_data, f, indent=4)
            except Exception:
                pass

            # Align DataFrame directly to our friendly column layout
            df = pd.DataFrame.from_dict(all_form_data, orient="index")
            df = df.reindex(columns=FINAL_COLUMNS)
            df.index.name = "Source File Name"
            df = df.fillna("")
            
            st.success(f"🎉 Successfully processed {len(all_form_data)} PDFs!")
            
            # --- VISUAL PREVIEW SECTION ---
            st.subheader("📊 Spreadsheet Preview")
            st.write("Review the extracted data below before downloading:")
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
