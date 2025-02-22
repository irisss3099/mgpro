import streamlit as st
import pandas as pd
import os
from io import BytesIO
from pdf2docx import Converter
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from PIL import Image

st.set_page_config(page_title="📚 File Transformer Pro 🚀", layout="wide")

# Custom CSS
st.markdown(
    """
    <style> 
    .stApp{
            background-color: black;
            color: white;
           }
           </style>
            """ ,
            unsafe_allow_html=True
)

# Title and Description
st.title("🔄 Ultimate File Transformer Hub 💼")
st.write("Convert files seamlessly between PDF, Word, Excel, PNG, CSV, PowerPoint, and more with advanced data cleaning and visualization capabilities.")

# File Uploader
uploaded_files = st.file_uploader("📤 Upload your files", type=["csv", "xlsx", "pdf", "png", "docx", "pptx"], accept_multiple_files=True)

if uploaded_files:
    for file in uploaded_files:
        file_ext = os.path.splitext(file.name)[-1].lower()

        # File Handling
        if file_ext == ".csv":
            df = pd.read_csv(file)
        elif file_ext == ".xlsx":
            df = pd.read_excel(file)
        else:
            df = None

        # File Preview (Only for CSV and Excel)
        if df is not None:
            st.write("📊 Preview the Data Frame")
            st.dataframe(df.head())

            # Data Cleaning Options
            st.subheader("🧹 Data Cleaning Options")
            if st.checkbox(f"Clean Data for {file.name}"):
                col1, col2 = st.columns(2)

                with col1:
                    if st.button(f"🗑️ Remove Duplicates from {file.name}"):
                        df.drop_duplicates(inplace=True)
                        st.write("✅ Duplicates Removed!")

                with col2:
                    if st.button(f"🧼 Fill Missing Values for {file.name}"):
                        numeric_cols = df.select_dtypes(include=['number']).columns
                        df[numeric_cols] = df[numeric_cols].fillna(df[numeric_cols].mean())
                        st.write("📈 Missing Values Filled!")

            # Column Selection
            st.header("📌 Select Columns to Keep")
            columns = st.multiselect(f"Choose Columns for {file.name}", df.columns, default=df.columns)
            df = df[columns]

            # Data Visualization
            st.subheader("📊 Data Visualization")
            if st.checkbox(f"Show Visualization for {file.name}"):
                st.bar_chart(df.select_dtypes(include='number').iloc[:, :2])

        # File Conversion
        st.subheader("🔄 File Conversion Options")
        conversion_type = st.radio(f"Convert {file.name} to:", ["CSV", "Excel", "PDF", "Word", "PNG", "PowerPoint"], key=file.name)

        if st.button(f"🚀 Convert {file.name}"):
            buffer = BytesIO()

            if conversion_type == "CSV" and df is not None:
                df.to_csv(buffer, index=False)
                file_name = file.name.replace(file_ext, ".csv")
                mime_type = "text/csv"

            elif conversion_type == "Excel" and df is not None:
                df.to_excel(buffer, index=False)
                file_name = file.name.replace(file_ext, ".xlsx")
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

            elif conversion_type == "Word" and file_ext == ".pdf":
                pdf_file_path = f"{file.name}"
                docx_file_path = pdf_file_path.replace(".pdf", ".docx")
                with open(pdf_file_path, "wb") as f:
                    f.write(file.getvalue())
                cv = Converter(pdf_file_path)
                cv.convert(docx_file_path, start=0, end=None)
                cv.close()
                with open(docx_file_path, "rb") as f:
                    buffer.write(f.read())
                file_name = docx_file_path
                mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

            elif conversion_type == "PDF" and file_ext == ".docx":
                doc = Document(file)
                pdf_path = file.name.replace(".docx", ".pdf")
                pdf = canvas.Canvas(buffer, pagesize=letter)
                pdf.drawString(100, 750, "Word to PDF Conversion")
                pdf.save()
                file_name = pdf_path
                mime_type = "application/pdf"

            elif conversion_type == "PNG" and file_ext == ".pdf":
                pdf_file_path = f"{file.name}"
                with open(pdf_file_path, "wb") as f:
                    f.write(file.getvalue())
                image = Image.open(pdf_file_path)
                png_path = pdf_file_path.replace(".pdf", ".png")
                image.save(png_path)
                with open(png_path, "rb") as f:
                    buffer.write(f.read())
                file_name = png_path
                mime_type = "image/png"

            elif conversion_type == "PDF" and file_ext == ".png":
                image = Image.open(file)
                pdf_path = file.name.replace(".png", ".pdf")
                image.save(buffer, "PDF")
                file_name = pdf_path
                mime_type = "application/pdf"

            buffer.seek(0)

            st.download_button(
                label=f"📥 Download {file.name} as {conversion_type}",
                data=buffer,
                file_name=file_name,
                mime=mime_type
            )

    st.success("🎉 All files processed successfully!")
