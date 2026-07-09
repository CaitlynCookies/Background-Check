import pandas as pd
import streamlit as st
from pypdf import PdfReader

st.title("PDF Form to Excel Converter")
st.write("Upload your PDFs, and we will extract the data into an Excel sheet!")

# 1. Create a file uploader on the webpage
uploaded_files = st.file_uploader(
    "Choose PDF files", type="pdf", accept_multiple_files=True
)

if uploaded_files:
    all_data = []

    # 2. Process the uploaded files
    for uploaded_file in uploaded_files:
        reader = PdfReader(uploaded_file)
        fields = reader.get_fields()

        if fields:
            file_data = {"File Name": uploaded_file.name}
            for field_name, field_data in fields.items():
                file_data[field_name] = field_data.get("/V", "")
            all_data.append(file_data)

    # 3. Show a preview table on the website
    if all_data:
        df = pd.DataFrame(all_data)
        st.subheader("Data Preview")
        st.dataframe(df)

        # 4. Create an Excel download button
        # (We save to system memory instead of a file path so web users can download it)
        import io

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)

        st.download_button(
            label="📥 Download Excel Sheet",
            data=buffer.getvalue(),
            file_name="extracted_pdf_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
