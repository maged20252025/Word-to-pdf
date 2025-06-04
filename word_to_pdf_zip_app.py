
import streamlit as st
import os
import tempfile
import zipfile
import shutil
from docx import Document
from pathlib import Path
from fpdf import FPDF

def convert_docx_to_pdf(docx_path, pdf_path):
    doc = Document(docx_path)
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Arial", size=12)
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            pdf.multi_cell(0, 10, text)
    pdf.output(pdf_path)

# إعداد الواجهة
st.set_page_config(page_title="📝 تحويل Word إلى PDF داخل ملف مضغوط", layout="centered")
st.title("📄 تحويل ملفات Word إلى PDF وتسليمها في ملف مضغوط")

uploaded_files = st.file_uploader(
    "📥 ارفع ملفات Word مباشرة أو داخل ملف ZIP",
    type=["doc", "docx", "zip"],
    accept_multiple_files=True
)

if uploaded_files:
    with st.spinner("🔄 جاري تحويل الملفات إلى PDF..."):
        temp_dir = tempfile.mkdtemp()
        input_dir = os.path.join(temp_dir, "input")
        output_dir = os.path.join(temp_dir, "pdfs")
        os.makedirs(input_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)

        # استخراج الملفات من ZIP أو حفظ الملفات المباشرة
        for uploaded in uploaded_files:
            name = uploaded.name.lower()
            if name.endswith(".zip"):
                zip_path = os.path.join(input_dir, uploaded.name)
                with open(zip_path, "wb") as f:
                    f.write(uploaded.read())
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_ref.extractall(input_dir)
            elif name.endswith((".doc", ".docx")):
                file_path = os.path.join(input_dir, uploaded.name)
                with open(file_path, "wb") as f:
                    f.write(uploaded.read())

        # تحويل ملفات Word إلى PDF
        word_files = []
        for root, _, files in os.walk(input_dir):
            for file in files:
                if file.lower().endswith(".docx"):
                    full_path = os.path.join(root, file)
                    word_files.append(full_path)

        if not word_files:
            st.error("❌ لم يتم العثور على ملفات .docx. الرجاء رفع ملفات Word بصيغة docx فقط.")
        else:
            for docx_file in word_files:
                file_name = Path(docx_file).stem + ".pdf"
                pdf_output_path = os.path.join(output_dir, file_name)
                convert_docx_to_pdf(docx_file, pdf_output_path)

            # ضغط ملفات PDF في ملف مضغوط
            final_zip = os.path.join(temp_dir, "ملفات_PDF.zip")
            with zipfile.ZipFile(final_zip, "w", zipfile.ZIP_DEFLATED) as zipf:
                for pdf_file in os.listdir(output_dir):
                    zipf.write(os.path.join(output_dir, pdf_file), pdf_file)

            with open(final_zip, "rb") as f:
                st.success(f"✅ تم تحويل {len(word_files)} ملف Word إلى PDF.")
                st.download_button(
                    label="📥 تحميل ملف PDF مضغوط",
                    data=f,
                    file_name="ملفات_PDF.zip",
                    mime="application/zip"
                )
