
import streamlit as st
import os
import tempfile
import zipfile
import shutil
from docx import Document
from pathlib import Path
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from bidi.algorithm import get_display
import arabic_reshaper

# تحميل خط عربي (يفترض أن يكون موجودًا في مجلد العمل أو يتم تحميله لاحقًا)
FONT_PATH = "arial.ttf"
pdfmetrics.registerFont(TTFont("Arabic", FONT_PATH))

def convert_docx_to_pdf(docx_path, pdf_path):
    doc = Document(docx_path)
    c = canvas.Canvas(pdf_path, pagesize=A4)
    width, height = A4
    c.setFont("Arabic", 14)
    y = height - 50

    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            reshaped_text = arabic_reshaper.reshape(text)
            bidi_text = get_display(reshaped_text)
            for line in bidi_text.split("\n"):
                c.drawRightString(width - 40, y, line)
                y -= 20
                if y < 50:
                    c.showPage()
                    c.setFont("Arabic", 14)
                    y = height - 50

    c.save()

# إعداد الواجهة
st.set_page_config(page_title="📝 تحويل Word إلى PDF داخل ملف مضغوط", layout="centered")
st.title("📄 تحويل ملفات Word (بالعربية) إلى PDF وتسليمها في ملف مضغوط")

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

        word_files = []
        for root, _, files in os.walk(input_dir):
            for file in files:
                if file.lower().endswith(".docx"):
                    full_path = os.path.join(root, file)
                    word_files.append(full_path)

        if not word_files:
            st.error("❌ لم يتم العثور على ملفات .docx. الرجاء رفع ملفات Word بصيغة docx فقط.")
        elif not os.path.exists(FONT_PATH):
            st.error("❌ لم يتم العثور على الخط العربي 'arial.ttf'. الرجاء رفعه مع الكود.")
        else:
            for docx_file in word_files:
                file_name = Path(docx_file).stem + ".pdf"
                pdf_output_path = os.path.join(output_dir, file_name)
                convert_docx_to_pdf(docx_file, pdf_output_path)

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
