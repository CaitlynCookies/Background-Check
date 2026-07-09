import json
from pathlib import Path
from pypdf import PdfReader

def extract_form_fields_batch(input_folder: str, output_json: str):
    all_form_data = {}
    pdf_dir = Path(input_folder)

    # Efficiently stream through the folder file-by-file
    for pdf_path in pdf_dir.glob("*.[pP][dD][fF]"):
        file_name = pdf_path.name

        try:
            reader = PdfReader(pdf_path)

            # Fast check: Skip the file entirely if it has no interactive form fields
            fields = reader.get_fields()
            if not fields:
                continue

            file_data = {}
            # Your exact snippet, optimized to store data instead of just printing it
            for field_name, field_data in fields.items():
                # Extract the value, default to an empty string if blank
                value = field_data.get("/V", "")

                # Clean up PDF formatting quirks (e.g., checkbox '/Yes' becomes 'Yes')
                if isinstance(value, str) and value.startswith("/"):
                    value = value.lstrip("/")

                file_data[field_name] = value

            all_form_data[file_name] = file_data

        except Exception as e:
            print(f"Error skipping/reading {file_name}: {e}")

    # Write all collected data to a JSON file at once
    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(all_form_data, f, indent=4, ensure_ascii=False)

extract_form_fields_batch("/content/BackgroundCheckTest", "extracted_forms.json")

import pandas as pd

with open("extracted_forms.json", "r", encoding="utf-8") as f:
    data = json.load(f)
df = pd.DataFrame.from_dict(data, orient="index")
headers = ["Personal_name_first~7", "Personal_name_last~7", "EF_Emp_phone_primary~7", "EF_Emp_birth_date~7", "EF_Emp_Residence_street_1~7", "EF_Emp_Residence_city~7", "EF_Emp_Residence_state_rdo~7", "EF_Emp_Residence_zip_code~7", "EF_Emp_ssn~7", "EF_Emp_email~7"]
df = df.reindex(columns = headers)
df.columns = ["First name", "Last name", "Phone", "DOB", "Address Street", "Address City", "Address State", "Address Zip", "SSN", "Email Address"]

df.to_excel("test.xlsx")
