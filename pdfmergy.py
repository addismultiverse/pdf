import streamlit as st
import PyPDF2
from PIL import Image
import os
import win32com.client  # Ensure pywin32 is installed for this
import tempfile

# PDF Merger with Streamlit
def merge_pdfs(file_paths, output_path):
    pdf_writer = PyPDF2.PdfWriter()
    
    for file_path in file_paths:
        ext = os.path.splitext(file_path)[1].lower()

        if ext == ".pdf":
            merge_pdf(file_path, pdf_writer)
        elif ext == ".docx":
            convert_word_to_pdf(file_path, pdf_writer)
        elif ext == ".xlsx":
            convert_excel_to_pdf(file_path, pdf_writer)
        elif ext in [".png", ".jpg", ".jpeg"]:
            convert_image_to_pdf(file_path, pdf_writer)

    with open(output_path, "wb") as output_file:
        pdf_writer.write(output_file)

def merge_pdf(file_path, pdf_writer):
    try:
        pdf_reader = PyPDF2.PdfReader(file_path)
        for page in range(len(pdf_reader.pages)):
            pdf_writer.add_page(pdf_reader.pages[page])
    except Exception as e:
        st.error(f"Error reading {file_path}: {e}")

def convert_word_to_pdf(file_path, pdf_writer):
    try:
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(file_path)
        temp_pdf = file_path.replace(".docx", ".pdf")
        doc.SaveAs(temp_pdf, FileFormat=17)  # 17 = PDF format
        doc.Close(False)
        word.Quit()
        merge_pdf(temp_pdf, pdf_writer)
        os.remove(temp_pdf)
    except Exception as e:
        st.error(f"Error converting Word document {file_path}: {e}")

def convert_excel_to_pdf(file_path, pdf_writer):
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        wb = excel.Workbooks.Open(file_path)
        temp_pdf = file_path.replace(".xlsx", ".pdf")
        wb.ExportAsFixedFormat(0, temp_pdf)  # 0 = PDF format
        wb.Close(False)
        excel.Quit()
        merge_pdf(temp_pdf, pdf_writer)
        os.remove(temp_pdf)
    except Exception as e:
        st.error(f"Error converting Excel document {file_path}: {e}")

def convert_image_to_pdf(file_path, pdf_writer):
    try:
        img = Image.open(file_path)
        temp_pdf = file_path.replace(os.path.splitext(file_path)[1], ".pdf")
        img.convert("RGB").save(temp_pdf)
        merge_pdf(temp_pdf, pdf_writer)
        os.remove(temp_pdf)
    except Exception as e:
        st.error(f"Error converting image {file_path}: {e}")

def main():
    st.title("Multi-Document Merger with Reorder Option")

    # Input fields for project, activity, MT, and lead trade
    project = st.text_input("Enter Project Number:")
    activity = st.text_input("Enter Activity Number:")
    mt = st.text_input("Enter MT Number:")
    lead_trade = st.text_input("Enter Lead Trade:")

    # File uploader (accepting multiple files)
    uploaded_files = st.file_uploader(
        "Select files (PDF, Word, Excel, Images):",
        type=["pdf", "docx", "xlsx", "png", "jpg", "jpeg"],
        accept_multiple_files=True
    )

    # Display uploaded files
    if uploaded_files:
        file_paths = []
        st.write("Uploaded Files:")
        for uploaded_file in uploaded_files:
            temp_file = tempfile.NamedTemporaryFile(delete=False)
            temp_file.write(uploaded_file.read())
            file_paths.append(temp_file.name)
            st.write(f"- {uploaded_file.name}")

        # Reorder and delete options (not fully implemented in Streamlit)
        st.warning("Reordering files is not implemented in this version.")

        if st.button("Merge Files"):
            if not project or not activity or not mt or not lead_trade:
                st.error("Please fill in all the fields.")
            else:
                output_filename = f"{project}-{activity}-{mt}-{lead_trade}.pdf"
                output_path = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False).name
                merge_pdfs(file_paths, output_path)

                # Provide the download link for the merged PDF
                with open(output_path, "rb") as file:
                    btn = st.download_button(
                        label="Download Merged PDF",
                        data=file,
                        file_name=output_filename,
                        mime="application/pdf"
                    )

if __name__ == "__main__":
    main()
