
import streamlit as st
import os
import tempfile
import zipfile
import shutil
from pathlib import Path
import subprocess

st.set_page_config(page_title="🔁 تحويل DOC إلى DOCX", layout="centered")
st.title("📄 تحويل ملفات Word القديمة (.doc) إلى Word الحديثة (.docx)")

uploaded_files = st.file_uploader(
    "📥 ارفع ملفات بصيغة .doc أو ملف ZIP يحتوي عليها",
    type=["doc", "zip"],
    accept_multiple_files=True
)

def convert_doc_to_docx_linux(input_path, output_path):
    # Requires LibreOffice installed
    command = ["libreoffice", "--headless", "--convert-to", "docx", "--outdir", os.path.dirname(output_path), input_path]
    subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

if uploaded_files:
    with st.spinner("🔄 جاري التحويل..."):
        temp_dir = tempfile.mkdtemp()
        input_dir = os.path.join(temp_dir, "input")
        output_dir = os.path.join(temp_dir, "docx_files")
        os.makedirs(input_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)

        for uploaded in uploaded_files:
            name = uploaded.name.lower()
            if name.endswith(".zip"):
                zip_path = os.path.join(input_dir, uploaded.name)
                with open(zip_path, "wb") as f:
                    f.write(uploaded.read())
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_ref.extractall(input_dir)
            elif name.endswith(".doc"):
                file_path = os.path.join(input_dir, uploaded.name)
                with open(file_path, "wb") as f:
                    f.write(uploaded.read())

        doc_files = []
        for root, _, files in os.walk(input_dir):
            for file in files:
                if file.lower().endswith(".doc"):
                    full_path = os.path.join(root, file)
                    doc_files.append(full_path)

        if not doc_files:
            st.error("❌ لم يتم العثور على ملفات .doc.")
        else:
            for doc_file in doc_files:
                convert_doc_to_docx_linux(doc_file, output_dir)

            for file in os.listdir(input_dir):
                if file.lower().endswith(".docx"):
                    shutil.move(os.path.join(input_dir, file), os.path.join(output_dir, file))

            final_zip = os.path.join(temp_dir, "ملفات_DOCX.zip")
            with zipfile.ZipFile(final_zip, "w", zipfile.ZIP_DEFLATED) as zipf:
                for file in os.listdir(output_dir):
                    zipf.write(os.path.join(output_dir, file), file)

            with open(final_zip, "rb") as f:
                st.success(f"✅ تم تحويل {len(doc_files)} ملف .doc إلى .docx.")
                st.download_button(
                    label="📥 تحميل الملفات المحوّلة",
                    data=f,
                    file_name="ملفات_DOCX.zip",
                    mime="application/zip"
                )
