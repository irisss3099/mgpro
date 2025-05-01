import streamlit as st
import pandas as pd
import os
from io import BytesIO
from docx import Document
from pypdf import PdfReader
from PIL import Image

# Set page config as the very first Streamlit command
st.set_page_config(page_title="File Converter Hub 📁", layout="wide")

# Custom CSS for styling
st.markdown("""
    <style>
    .stApp {
        background-color: black;
        color: white;
    }
    </style>
    """, unsafe_allow_html=True)

# Title & Description
st.title("File Converter Hub 📁")
st.write("Easily convert and transform your files between multiple formats including CSV, Excel, PDF, Word, and Images!")

# File Uploader
uploaded_files = st.file_uploader("💾 Upload your files", type=["csv", "xlsx", "pdf", "docx", "png", "jpg", "jpeg"], accept_multiple_files=True)

if uploaded_files:
    for file in uploaded_files:
        file_ext = os.path.splitext(file.name)[-1].lower()

        # Initialize variables
        df = None
        pdf_text = ""
        image = None

        try:
            # File reading
            if file_ext == ".csv":
                df = pd.read_csv(file)
            elif file_ext == ".xlsx":
                df = pd.read_excel(file)
            elif file_ext == ".pdf":
                reader = PdfReader(file)
                pdf_text = "\n".join([page.extract_text() for page in reader.pages])
            elif file_ext == ".docx":
                doc = Document(file)
                pdf_text = "\n".join([para.text for para in doc.paragraphs])
            elif file_ext in [".png", ".jpg", ".jpeg"]:
                image = Image.open(file)
            else:
                st.error(f"⛔ Unsupported file type: {file_ext}")
                continue
        except Exception as e:
            st.error(f"❌ Error reading file: {str(e)}")
            continue

        # Preview the file content
        st.write(f"📄 Preview of {file.name}")
        if df is not None:
            st.dataframe(df.head())
        elif pdf_text:
            st.text_area("Extracted Text", pdf_text, height=300)
        elif image:
            st.image(image, caption=file.name, use_container_width=True)  # Changed here

        # File Conversion Options
        st.subheader("🔄 Conversion Options")
        conversion_types = ["CSV", "Excel", "PDF", "Word", "PNG", "JPG", "JPEG"]
        conversion_type = st.radio(f"Convert {file.name} to:", conversion_types, key=file.name)

        if st.button(f"Convert {file.name}"):
            buffer = BytesIO()

            try:
                # Handle conversion logic based on selected file type and destination
                if conversion_type == "CSV":
                    if df is not None:
                        df.to_csv(buffer, index=False)
                        file_name = file.name.replace(file_ext, ".csv")
                        mime_type = "text/csv"
                    else:
                        st.warning("⚠️ Only CSV files can be converted to CSV.")
                        continue

                elif conversion_type == "Excel":
                    if df is not None:
                        df.to_excel(buffer, index=False)
                        file_name = file.name.replace(file_ext, ".xlsx")
                        mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    else:
                        st.warning("⚠️ Only Excel files can be converted to Excel.")
                        continue

                elif conversion_type == "PDF":
                    if pdf_text:
                        buffer.write(pdf_text.encode("utf-8"))
                        file_name = file.name.replace(file_ext, ".pdf")
                        mime_type = "application/pdf"
                    else:
                        st.warning("⚠️ Only PDF text can be converted to PDF.")
                        continue

                elif conversion_type == "Word":
                    if pdf_text:
                        doc = Document()
                        for line in pdf_text.split("\n"):
                            doc.add_paragraph(line)
                        doc.save(buffer)
                        file_name = file.name.replace(file_ext, ".docx")
                        mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    else:
                        st.warning("⚠️ Only Word text can be converted to Word.")
                        continue

                elif conversion_type == "PNG":
                    if image:
                        image.save(buffer, format="PNG")
                        file_name = file.name.replace(file_ext, ".png")
                        mime_type = "image/png"
                    else:
                        st.warning("⚠️ Only image files can be converted to PNG.")
                        continue

                elif conversion_type == "JPG" or conversion_type == "JPEG":
                    if image:
                        image.save(buffer, format="JPEG")
                        file_name = file.name.replace(file_ext, ".jpg")
                        mime_type = "image/jpeg"
                    else:
                        st.warning("⚠️ Only image files can be converted to JPG/JPEG.")
                        continue

                else:
                    st.warning("⚠️ Unsupported or missing data for conversion.")
                    continue

                # Prepare buffer for download
                buffer.seek(0)

                st.download_button(
                    label=f"📥 Download {file.name} as {conversion_type}",
                    data=buffer,
                    file_name=file_name,
                    mime=mime_type
                )

            except Exception as e:
                st.error(f"❌ Conversion failed: {str(e)}")

    st.success("✅ All files processed successfully!")



  
