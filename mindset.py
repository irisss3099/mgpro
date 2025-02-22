from typing_extensions import Doc
import streamlit as st
import pandas as pd
import os
from io import BytesIO
import PDF2
from PIL import Image
Doc
import Document 

st.set_page_config(page_title="File Converter Hub 📁", layout="wide")

# Custom CSS
st.markdown(
    """
    <style>
    .stApp {
        background-color: black;
        color: white;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Title and Description
st.title("File Converter Hub 📁")
st.write("Easily convert and transform your files between multiple formats including CSV, Excel, PDF, Word, and Images!")

# File Uploader
uploaded_files = st.file_uploader("💾Upload your files", 
                                  type=["csv", "xlsx", "pdf", "docx", "png", "jpg"], 
                                  accept_multiple_files=True)

if uploaded_files:
    for file in uploaded_files:
        file_ext = os.path.splitext(file.name)[-1].lower()

        # Read Files
        if file_ext == ".csv":
            df = pd.read_csv(file)
        elif file_ext == ".xlsx":
            df = pd.read_excel(file)
        elif file_ext == ".pdf":
            pdf_reader = pdf_reader(file)
            pdf_text = "\n".join(page.extract_text() for page in pdf_reader.pages)
        elif file_ext == ".docx":
            doc = Document(file)
            pdf_text = "\n".join([para.text for para in doc.paragraphs])
        elif file_ext in [".png", ".jpg"]:
            image = Image.open(file)
        else:
            st.error(f"⛔ Unsupported file type: {file_ext}")
            continue

        # File Preview
        st.write(f"📄  Preview of {file.name}")
        if file_ext in [".csv", ".xlsx"]:
            st.dataframe(df.head())
        elif file_ext in [".pdf", ".docx"]:
            st.text_area("Extracted Text", pdf_text, height=300)
        elif file_ext in [".png", ".jpg"]:
            st.image(image, caption=file.name, use_column_width=True)

        # Data Cleaning Options (For CSV & Excel)
        if file_ext in [".csv", ".xlsx"]:
            st.subheader("🧹 Data Cleaning Options")
            if st.checkbox(f"Clean Data for {file.name}"):
                col1, col2 = st.columns(2)

                with col1:
                    if st.button(f"🗑️ Remove Duplicates from {file.name}"):
                        df.drop_duplicates(inplace=True)
                        st.write("✅ Duplicates Removed!")

                with col2:
                    if st.button(f"🧹 Fill Missing Values for {file.name}"):
                        numeric_cols = df.select_dtypes(include=['number']).columns
                        df[numeric_cols] = df[numeric_cols].fillna(df[numeric_cols].mean())
                        st.write("✅ Missing Values Filled!")

            st.header("📊 Select Columns to Keep")
            columns = st.multiselect(f"Choose Columns for {file.name}", df.columns, default=df.columns)
            df = df[columns]

            # Data Visualization
            st.subheader("📈 Data Visualization")
            if st.checkbox(f"Show Visualization for {file.name}"):
                st.bar_chart(df.select_dtypes(include='number').iloc[:, :2])

        # Conversion Options
        st.subheader("🔄 Conversion Options")
        conversion_types = ["CSV", "Excel", "PDF", "Word"]

        if file_ext in [".png", ".jpg"]:
            conversion_types.append("PNG")

        conversion_type = st.radio(f"Convert {file.name} to:", conversion_types, key=file.name)

        if st.button(f"Convert {file.name}"):
            buffer = BytesIO()

            if conversion_type == "CSV" and file_ext != ".csv":
                df.to_csv(buffer, index=False)
                file_name = file.name.replace(file_ext, ".csv")
                mime_type = "text/csv"

            elif conversion_type == "Excel" and file_ext != ".xlsx":
                df.to_excel(buffer, index=False)
                file_name = file.name.replace(file_ext, ".xlsx")
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

            elif conversion_type == "PDF" and file_ext not in [".pdf"]:
                buffer.write(pdf_text.encode("utf-8"))
                file_name = file.name.replace(file_ext, ".pdf")
                mime_type = "application/pdf"

            elif conversion_type == "Word" and file_ext not in [".docx"]:
                doc = Document()
                for line in pdf_text.split("\n"):
                    doc.add_paragraph(line)
                doc.save(buffer)
                file_name = file.name.replace(file_ext, ".docx")
                mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

            elif conversion_type == "PNG" and file_ext not in [".png"]:
                image.save(buffer, format="PNG")
                file_name = file.name.replace(file_ext, ".png")
                mime_type = "image/png"

            buffer.seek(0)

            st.download_button(
                label=f"📥 Download {file.name} as {conversion_type}",
                data=buffer,
                file_name=file_name,
                mime=mime_type
            )

    st.success("✅ All files processed successfully!")


  
