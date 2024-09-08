import streamlit as st
import PyPDF2
import os

# Set up the upload folder
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Streamlit app
st.title("PDF Merger")

# Input fields
project = st.text_input("Enter Project Number:")
activity = st.text_input("Enter Activity Number:")
mt = st.text_input("Enter MT Number:")
lead_trade = st.text_input("Enter Lead Trade:")

# File upload fields
file1 = st.file_uploader("Select first PDF file:", type="pdf")
file2 = st.file_uploader("Select second PDF file:", type="pdf")

if st.button("Merge PDFs"):
    if not project or not activity or not mt or not lead_trade:
        st.error("Please fill in all the fields.")
    elif not file1 or not file2:
        st.error("Please upload both PDF files.")
    else:
        try:
            file1_path = os.path.join(UPLOAD_FOLDER, file1.name)
            file2_path = os.path.join(UPLOAD_FOLDER, file2.name)
            
            # Save uploaded files to the upload folder
            with open(file1_path, "wb") as f:
                f.write(file1.getbuffer())
            with open(file2_path, "wb") as f:
                f.write(file2.getbuffer())

            output_filename = f"{project}-{activity}-{mt}-{lead_trade}.pdf"
            output_path = os.path.join(UPLOAD_FOLDER, output_filename)

            # Merge PDFs
            pdf_writer = PyPDF2.PdfWriter()
            for file_path in [file1_path, file2_path]:
                pdf_reader = PyPDF2.PdfReader(file_path)
                for page in range(len(pdf_reader.pages)):
                    pdf_writer.add_page(pdf_reader.pages[page])

            with open(output_path, "wb") as output_file:
                pdf_writer.write(output_file)

            with open(output_path, "rb") as f:
                st.success("PDF files have been merged successfully!")
                st.download_button(
                    label="Download Merged PDF",
                    data=f,
                    file_name=output_filename,
                    mime="application/pdf"
                )

        except Exception as e:
            st.error(f"An error occurred: {e}")
