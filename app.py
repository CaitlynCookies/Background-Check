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
    
    # --- DATE OF BIRTH ---
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
    
    # --- PHONE ---
    "EF_Emp_phone_primary~3": "Phone",
    "EF_Emp_phone_primary~4": "Phone",
    "EF_Emp_phone_primary~7": "Phone",
    
    # --- EMAIL ---
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
    "Employee_W4_qualifying_deps_amount~7": "W4 Qualifying
