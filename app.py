import streamlit as st
from pdf2docx import Converter
from fpdf import FPDF
from docx import Document


# Apply custom styling
st.markdown(
    """
    <style>
        .main {
            background: #d3d3d3;
            padding: 20px;
            border-radius: 15px;
            text-align: center;
            color: black;
        }
        .stButton>button {
            background-color: #4CAF50;
            color: white;
            padding: 10px 24px;
            border: none;
            border-radius: 10px;
            cursor: pointer;
            font-size: 16px;
        }
        .stButton>button:hover {
            background-color: #45a049;
        }
    </style>
    """,
    unsafe_allow_html=True
)

# Colorful Header
st.markdown('<div class="main"><h1>PDF to Word & Word to PDF Converter</h1><h3>Created by Kaustav Roy Chowdhury</h3></div>', unsafe_allow_html=True)

# Upload File

uploaded_file = st.file_uploader("Upload a PDF or Word File", type=["pdf", "docx"])

if uploaded_file:
    file_name = uploaded_file.name
    file_ext = file_name.split(".")[-1].lower()

    # **PDF to Word Conversion**

    if file_ext == "pdf":
        st.subheader("Convert PDF to Word")

        if st.button("Convert"):
            pdf_path = "input.pdf"
            docx_path = "converted.pdf"

            # Save the uploaded PDF
            with open(pdf_path, "wb") as pdf_file:
                pdf_file.write(uploaded_file.read())


            # Convert PDF to Word

            cv = Converter(pdf_path)
            cv.convert(docx_path)
            cv.close()

            # Provide download link

            with open(docx_path, "rb") as docx_file:
                st.download_button("Download Word File", docx_file, file_name= "converted.docx")



     # **Word to PDF Conversion (Fixed Formatting)**
    elif file_ext == "docx":
        st.subheader("Convert Word to PDF")
        if st.button("Convert"):
            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)  # Auto break when text reaches bottom
            pdf.add_page()
            pdf.set_font("Arial", size=12)

            # Read Word file
            doc = Document(uploaded_file)
            for para in doc.paragraphs:
                pdf.multi_cell(0, 10, para.text)  # Use multi_cell for text wrapping
                pdf.ln(5)  # Add extra spacing

            pdf_path = "converted.pdf"
            pdf.output(pdf_path)


            # Provide download link

            with open(pdf_path, "rb") as pdf_file:
                st.download_button("Download PDF File", pdf_file, file_name="converted.pdf")


            


